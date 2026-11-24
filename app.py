"""Confirmaciones Paraty: Gmail / .eml -> confirmación Word.

Gmail OAuth local (solo lectura); también admite cargar .eml como respaldo.
La imagen opcional debe estar enmascarada: jamás pegar PAN completo/CVV.
La imagen solo vive en la sesión y en el Word que el operador decide descargar.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from email.header import decode_header, make_header
import base64
import os
from decimal import Decimal, ROUND_HALF_UP
from email import policy
from email.parser import BytesParser
from io import BytesIO
from html import escape as html_escape
from PIL import Image, ImageOps
import re
import unicodedata

from bs4 import BeautifulSoup
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

# Impuesto definido por el hotel para los servicios adicionales. No confundir
# con la tasa de impuesto sobre el alojamiento que informa Paraty.
ADDITIONAL_SERVICES_TAX_RATE = Decimal("0.16")  # Solo para compatibilidad/pruebas
# EDITA ESTA LÍNEA para cambiar el tipo de cambio usado por American Express.
DEFAULT_EXCHANGE_RATE = Decimal("18.20")


@dataclass
class Night:
    date: str
    original: Decimal
    discount: Decimal
    net: Decimal


@dataclass
class RoomDetail:
    number: int
    name: str
    adults: int = 0
    children: int = 0
    infants: int = 0
    night_rates: list[Night] = field(default_factory=list)
    stated_total: Decimal | None = None

    @property
    def total(self) -> Decimal:
        return sum((n.net for n in self.night_rates), Decimal("0"))

    @property
    def first_night(self) -> Decimal | None:
        return self.night_rates[0].net if self.night_rates else None


@dataclass
class AdditionalService:
    name: str
    quantity: int
    date: str
    price: Decimal
    currency: str = "USD"


@dataclass
class Booking:
    reference: str = ""
    origin: str = ""
    device: str = ""
    name: str = ""
    email: str = ""
    phone: str = ""
    location: str = ""
    requirements: str = ""
    arrival_time: str = ""
    total: Decimal = Decimal("0")
    taxes: Decimal = Decimal("0")
    pending: Decimal = Decimal("0")
    currency: str = "USD"
    checkin: str = ""
    checkout: str = ""
    nights: int = 0
    adults: int = 0
    children: int = 0
    room: str = ""
    meal_plan: str = ""
    rate: str = ""
    cancellation_policy: str = ""
    promotion: str = ""
    comments: str = ""
    benefits: list[str] = field(default_factory=list)
    terms: str = ""
    night_rates: list[Night] = field(default_factory=list)
    rooms: list[RoomDetail] = field(default_factory=list)
    deposit_required: Decimal | None = None
    deposit_room_average: Decimal | None = None
    deposit_services_charge: Decimal = Decimal("0")
    deposit_base_amount: Decimal | None = None
    flexible_rate: bool = False
    policy_status: str = "revisar"
    policy_reason: str = ""
    deposit_tax_included_verified: bool = False
    additional_services: list[AdditionalService] = field(default_factory=list)
    room_total: Decimal = Decimal("0")
    services_total: Decimal = Decimal("0")  # Base de servicios antes de impuesto
    services_tax_rate: Decimal | None = None  # Porcentaje elegido por el operador
    services_tax: Decimal = Decimal("0")
    services_with_tax: Decimal = Decimal("0")
    extra_difference: Decimal = Decimal("0")  # Total Paraty menos habitaciones y servicios con IVA
    services_tax_reconciled: bool = False
    tax_rate: Decimal | None = None
    warnings: list[str] = field(default_factory=list)


def money(raw: str) -> Decimal:
    hit = re.search(r"-?\d[\d,]*(?:\.\d+)?", raw.replace("\xa0", " "))
    return Decimal(hit.group(0).replace(",", "")) if hit else Decimal("0")


def field_value(soup: BeautifulSoup, label: str) -> str:
    # Paraty incluye etiquetas <b>Campo:</b> valor en el mismo nodo padre.
    for el in soup.find_all(["b", "strong"]):
        if el.get_text(" ", strip=True).strip().rstrip(":").strip().casefold() == label.casefold():
            # Evitar que textos de nodos hermanos distantes se mezclen con el campo.
            result = []
            for sibling in el.next_siblings:
                if getattr(sibling, "name", None) == "br":
                    break
                text = sibling.get_text(" ", strip=True) if hasattr(sibling, "get_text") else str(sibling)
                if text.strip():
                    result.append(text.strip())
            return re.sub(r"\s+", " ", " ".join(result)).strip()
    return ""


def after_label_text(soup: BeautifulSoup, label: str, stop_labels: list[str]) -> str:
    content = field_value(soup, label)
    if content:
        return content
    return ""



def classify_from_cancellation_terms(booking: Booking) -> tuple[str, str]:
    """Clasifica por políticas de cancelación y pago; NO por el nombre Rate.

    Nunca asumir depósito de una noche sin evidencia explícita de pago.
    Si la política es ambigua o contradictoria, requerir revisión.
    """
    policy_text = " ".join([booking.cancellation_policy, booking.terms]).strip()
    normalized = unicodedata.normalize("NFKD", policy_text).encode("ascii", "ignore").decode().lower()
    normalized = re.sub(r"\s+", " ", normalized)
    if not normalized:
        return "revisar", "No aparecen condiciones de cancelación o pago verificables."

    free_cancellation = bool(re.search(
        r"free cancell?ation|free of charge.{0,90}cancel|cancel.{0,100}(?:for free|free of charge)|"
        r"cancel.{0,90}(?:without (?:charge|penalty)|sin cargo|sin costo|sin penalizacion|gratuita?)|"
        r"cancelacion gratuita", normalized
    ))
    no_cancellation = bool(re.search(
        r"cancel(?:lations?|ations?|ing)?(?: and modifications)? (?:are |is )?"
        r"(?:not permitted|not allowed|prohibited)|"
        r"cannot be cancel(?:led|ed)|no se (?:permite|permiten) (?:la |las )?cancelaciones?|"
        r"cancelaciones? no (?:permitidas?|admitidas?)", normalized
    ))
    nonrefundable = bool(re.search(
        r"non[ -]?refundable|no reembolsable|"
        r"(?:if you cancel|si cancelas|si cancela).{0,130}no refund", normalized
    ))
    one_night_deposit = bool(re.search(
        r"(?:deposit|deposito|anticipo).{0,90}?(?:1|one|una?)\s*(?:night|noche)|"
        r"(?:1|one|una?)\s*(?:night|noche).{0,90}?(?:deposit|deposito|anticipo)",
        normalized
    ))
    full_payment = bool(re.search(
        r"(?:payment|pay|pago|cobro).{0,60}?100\s*%|"
        r"100\s*%.{0,90}?(?:time of booking|time of reservation|al reservar|reserva)|"
        r"(?:full|total) (?:payment|amount) (?:due|required) (?:at|upon) booking", normalized
    ))

    if free_cancellation and (no_cancellation or nonrefundable):
        return "revisar", "Hay condiciones contradictorias de cancelación gratuita y no reembolsable."
    if no_cancellation or nonrefundable:
        if one_night_deposit:
            return "revisar", "La política restringe cancelaciones, pero también menciona depósito de una noche."
        if not full_payment:
            return "revisar", "Cancelación no reembolsable, pero no se especifica cobro total anticipado."
        return "no_reembolsable", "Cancelaciones no permitidas o sin reembolso y pago total anticipado."
    if free_cancellation and one_night_deposit:
        if full_payment:
            return "revisar", "La política indica depósito de una noche y también pago total anticipado."
        return "flexible", "Cancelación gratuita según plazo y depósito explícito de una noche."
    if free_cancellation:
        return "revisar", "Se permite cancelación gratuita, pero no se confirma depósito de una noche."
    if one_night_deposit:
        return "revisar", "Se menciona depósito de una noche, pero la política de cancelación no es clara."
    return "revisar", "No se identifica una política de depósito y cancelación concluyente."


def policy_label(booking: Booking) -> str:
    return {"flexible": "Flexible - depósito de una noche",
            "no_reembolsable": "No reembolsable",
            "revisar": "REVISAR POLÍTICAS ANTES DEL COBRO"}.get(booking.policy_status, "Revisar")


def tax_percent_label(r: Booking) -> str:
    """Etiqueta humana para la tasa elegida (p.ej. 16% o 30%)."""
    if r.services_tax_rate is None:
        return "sin seleccionar"
    value = format(r.services_tax_rate * Decimal("100"), "f")
    if "." in value:
        value = value.rstrip("0").rstrip(".")
    return f"{value}%"


def deposit_service_label(r: Booking) -> str:
    """Texto apropiado tanto con suplementos como sin ellos."""
    return (f"servicios con impuesto {tax_percent_label(r)}"
            if r.additional_services else "sin servicios adicionales")


def recalculate_booking_finances(r: Booking, service_tax_rate: Decimal | None) -> None:
    """Recalcula servicio/impuestos y depósito a partir de una tasa ELEGIDA.

    La tarifa nocturna de cada habitación ya incluye los impuestos de Paraty.
    El servicio figura con precio base; el impuesto seleccionado se aplica una
    sola vez. No se modifica el total oficial de Paraty ni se cobran diferencias
    no conciliadas automáticamente.
    """
    r.warnings = [w for w in r.warnings
                  if not w.startswith(("IMPUESTOS SERVICIOS:", "DEPÓSITO:"))]
    r.deposit_required = None
    r.deposit_room_average = None
    r.deposit_base_amount = None
    r.deposit_services_charge = Decimal("0")
    r.deposit_tax_included_verified = False
    r.services_tax = Decimal("0")
    r.services_with_tax = Decimal("0")
    r.extra_difference = Decimal("0")
    r.services_tax_reconciled = False
    r.services_tax_rate = service_tax_rate if r.additional_services else None

    if r.additional_services:
        if service_tax_rate is None:
            r.warnings.append("IMPUESTOS SERVICIOS: selecciona el impuesto aplicable "
                              "(16%, 30% u otro) antes de generar la confirmación.")
            return
        if not Decimal("0") <= service_tax_rate <= Decimal("1"):
            raise ValueError("La tasa de impuesto de servicios debe estar entre 0% y 100%.")
        # Redondeo por servicio (tal como se presenta en el documento).
        r.services_tax = sum(
            ((item.price * service_tax_rate).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
             for item in r.additional_services), Decimal("0")
        )
        r.services_with_tax = r.services_total + r.services_tax
        if r.night_rates:
            r.extra_difference = (r.total - r.room_total - r.services_with_tax).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            r.services_tax_reconciled = abs(r.extra_difference) <= Decimal("0.02")
            if not r.services_tax_reconciled:
                r.warnings.append(
                    "IMPUESTOS SERVICIOS: el total de Paraty no coincide con "
                    f"habitaciones + suplementos + impuesto {tax_percent_label(r)} "
                    f"(diferencia: {fmt_usd(r.extra_difference)}). "
                    "Verifica los importes antes de cobrar; la diferencia no "
                    "se añade automáticamente al depósito."
                )

    if not r.flexible_rate:
        return
    if r.nights <= 0 or r.total <= 0:
        r.warnings.append("DEPÓSITO: faltan noches o importe total. Revisar Paraty.")
    elif not r.rooms or len(r.rooms) > 4:
        r.warnings.append("DEPÓSITO: deben identificarse entre 1 y 4 habitaciones.")
    elif any(len(room.night_rates) != r.nights for room in r.rooms):
        r.warnings.append("DEPÓSITO: falta un desglose completo por habitación y noche.")
    else:
        # Con suplementos: prorratear solo habitaciones y añadir el servicio con
        # el impuesto elegido completo al primer cargo. Sin suplementos: prorratear
        # el importe total oficial, tal como se hacía anteriormente.
        room_amount = r.room_total if r.additional_services else r.total
        if room_amount <= 0:
            r.warnings.append("DEPÓSITO: el subtotal de alojamiento no es válido.")
            return
        r.deposit_base_amount = room_amount
        r.deposit_room_average = (room_amount / Decimal(r.nights)).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        r.deposit_services_charge = r.services_with_tax
        r.deposit_required = (r.deposit_room_average + r.deposit_services_charge).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        r.deposit_tax_included_verified = (not r.additional_services or r.services_tax_reconciled)


def parse_paraty_eml(data: bytes) -> Booking:
    msg = BytesParser(policy=policy.default).parsebytes(data)
    htmlpart = msg.get_body(preferencelist=("html", "plain"))
    if htmlpart is None:
        raise ValueError("Este correo no tiene contenido HTML o de texto.")
    soup = BeautifulSoup(htmlpart.get_content(), "html.parser")
    ref_el = soup.find(attrs={"itemprop": "reservationNumber"})
    ref = (ref_el.get("content") if ref_el else None) or field_value(soup, "Booking reference")
    if not ref:
        raise ValueError("No se encontró una referencia de reservación Paraty en el correo.")
    r = Booking(reference=ref)
    for key, lab in [
        ("origin", "Origin"), ("device", "Device"), ("name", "Name"),
        ("email", "Email address"), ("phone", "Telephone"),
        ("requirements", "Special requirements"),
        ("arrival_time", "Estimated Time of Arrival at the Hotel"),
        ("room", "Room Selected"), ("meal_plan", "Selected Board Plan"),
        ("rate", "Rate"), ("promotion", "Promotion"),
    ]:
        setattr(r, key, field_value(soup, lab))
    # En Paraty, la política de cancelación flexible aparece en un <span>
    # inmediatamente después de <b>Cancellation policy:</b>.
    cp_label = soup.find(["b", "strong"], string=lambda t: t and "Cancellation policy" in t)
    if cp_label:
        for next_node in cp_label.next_siblings:
            if getattr(next_node, "name", None) == "span":
                r.cancellation_policy = next_node.get_text(" ", strip=True)
                break
    r.location = field_value(soup, "Address").lstrip(": ").strip()
    r.total = money(field_value(soup, "Total price"))
    # En Paraty «Pending payment» es texto del <span>, no una etiqueta <b>.
    pending_el = soup.find(string=lambda z: z and "Pending payment:" in z)
    if pending_el:
        r.pending = money(pending_el.parent.get_text(" ", strip=True).split(":", 1)[-1])
    tax_b = soup.find(["b", "strong"], string=lambda t: t and "Taxes per stay" in t)
    if tax_b:
        r.taxes = money(tax_b.parent.get_text(" ", strip=True).split(":", 1)[-1])
        rate_hit = re.search(r"(\d+(?:\.\d+)?)\s*%", tax_b.get_text(" ", strip=True))
        if rate_hit:
            r.tax_rate = Decimal(rate_hit.group(1)) / Decimal("100")
    for attr, item in [("checkin", "checkinDate"), ("checkout", "checkoutDate")]:
        el = soup.find(attrs={"itemprop": item})
        if el and el.get("content"):
            setattr(r, attr, el.get("content")[:10])
    if r.checkin and r.checkout:
        r.nights = (datetime.fromisoformat(r.checkout) - datetime.fromisoformat(r.checkin)).days
    r.adults = int((re.search(r"(\d+)\s+Adults", soup.get_text(" ", strip=True)) or [None, 0])[1])
    r.children = int((re.search(r"(\d+)\s+Children", soup.get_text(" ", strip=True)) or [None, 0])[1])
    # Cada habitación es una tabla independiente. Paraty puede enviar 1 a 4
    # (o más) tablas dentro del mismo bloque. Nunca interpretar solo la primera.
    for index, old in enumerate(soup.select(".old_price_per_day_table > table"), start=1):
        title_cell = old.select_one("tr > td > b")
        title = title_cell.get_text(" ", strip=True) if title_cell else ""
        meta = re.search(r"Room\s*(\d+)\s*:\s*(.*?)\s*\((\d+)-(\d+)-(\d+)\)", title, re.I)
        if meta:
            number = int(meta.group(1))
            room = RoomDetail(number, meta.group(2).strip(), int(meta.group(3)),
                              int(meta.group(4)), int(meta.group(5)))
        else:
            name = re.sub(r"^Room\s*\d+\s*:\s*", "", title, flags=re.I).strip()
            room = RoomDetail(index, name or r.room or "Habitación no identificada")
            r.warnings.append(f"Revisar categoría u ocupación de la habitación {index}")

        outer_rows = old.find_all("tr", recursive=False)
        if len(outer_rows) >= 2:
            for col in outer_rows[1].find_all("td", recursive=False):
                nested = col.find("table")
                if not nested:
                    continue
                trs = nested.find_all("tr")
                if len(trs) < 2:
                    continue
                date_str = trs[0].get_text(" ", strip=True)
                amounts = re.findall(r"-?\s*\d+(?:[.,]\d+)?\s*USD", trs[1].get_text(" ", strip=True), re.I)
                if len(amounts) >= 3:
                    room.night_rates.append(Night(
                        date_str, money(amounts[0]), abs(money(amounts[1])), money(amounts[2])
                    ))
        total_text = outer_rows[-1].get_text(" ", strip=True) if outer_rows else ""
        total_hit = re.search(r"\bTotal\s*:\s*([\d,.]+)", total_text, re.I)
        if total_hit:
            room.stated_total = money(total_hit.group(1))
        r.rooms.append(room)
        r.night_rates.extend(room.night_rates)  # Compatibilidad con uso anterior.

    if r.rooms:
        r.adults = sum(x.adults for x in r.rooms)
        r.children = sum(x.children for x in r.rooms)
    # Paraty añade estos servicios fuera de la tabla de tarifas por noche.
    # "Price" es el importe mostrado para la línea, no se vuelve a multiplicar
    # por Quantity porque el correo no afirma que sea tarifa unitaria.
    for element in soup.select(".additional_service_element"):
        name_el = element.select_one("div > b")
        name = name_el.get_text(" ", strip=True) if name_el else ""
        content = re.sub(r"\s+", " ", element.get_text(" ", strip=True))
        qty_hit = re.search(r"Quantity\s*:\s*(\d+)", content, re.I)
        date_hit = re.search(r"Date of service\s*:\s*(.+?)(?=\s+Price\s*:|$)", content, re.I)
        price_el = element.select_one(".value_elem.currencyValue")
        curr_el = element.select_one(".monedaConv")
        price_hit = re.search(r"Price\s*:\s*([\d,.]+)", content, re.I)
        price = money(price_el.get_text(" ", strip=True) if price_el else (price_hit.group(1) if price_hit else ""))
        currency = curr_el.get_text(" ", strip=True) if curr_el else r.currency
        if not name:
            name = content.split("Quantity:", 1)[0].strip() or "Servicio adicional"
        if price_el is None and price_hit is None:
            r.warnings.append(f"No se identificó el importe del servicio adicional: {name}")
        if currency.upper() != r.currency.upper():
            r.warnings.append(f"Revisar moneda del servicio '{name}': {currency}")
        r.additional_services.append(AdditionalService(
            name=name,
            quantity=int(qty_hit.group(1)) if qty_hit else 1,
            date=re.sub(r"\s+,", ",", date_hit.group(1).strip()) if date_hit else "",
            price=price,
            currency=currency,
        ))

    r.room_total = sum((room.total for room in r.rooms), Decimal("0"))
    r.services_total = sum((x.price for x in r.additional_services), Decimal("0"))

    # La categoría de cobro se define solo DESPUÉS de leer las condiciones
    # de cancelación y pago. El nombre comercial de la tarifa es informativo.
    # La extracción de BOOKING TERMS AND CONDITIONS se hace antes del cálculo.
    terms_b = soup.find(["b", "strong"], string=lambda t: t and "BOOKING TERMS AND CONDITIONS" in t)
    if terms_b:
        parts = []
        for sib in terms_b.next_siblings:
            if getattr(sib, "name", None) in ["div", "p"]:
                parts.append(sib.get_text(" ", strip=True))
            elif isinstance(sib, str):
                parts.append(sib.strip())
        r.terms = re.sub(r"\s+", " ", " ".join(x for x in parts if x))
    r.policy_status, r.policy_reason = classify_from_cancellation_terms(r)
    # Si el título comercial contradice frontalmente las políticas, no
    # automatizamos un cobro aunque existan términos interpretables.
    label = r.rate.lower()
    if (r.policy_status == "flexible" and
            re.search(r"non[ -]?refundable|\bnrf\b|no reembolsable", label)):
        r.policy_status = "revisar"
        r.policy_reason = ("El nombre de la tarifa indica no reembolsable, "
                           "pero las políticas describen cancelación gratuita y depósito.")
    elif (r.policy_status == "no_reembolsable" and "flexible" in label):
        r.policy_status = "revisar"
        r.policy_reason = ("El nombre de la tarifa dice flexible, pero las "
                           "políticas indican cancelación no reembolsable.")
    r.flexible_rate = r.policy_status == "flexible"
    if r.policy_status == "revisar":
        r.warnings.append("POLÍTICA POR REVISAR: " + r.policy_reason +
                          " No se calcula ni recomienda un depósito automático.")

    # Recálculo inicial con la tasa heredada. En Streamlit se vuelve a calcular
    # con el porcentaje que el operador elija para los servicios adicionales.
    recalculate_booking_finances(r, ADDITIONAL_SERVICES_TAX_RATE)

    # Beneficios: lista bajo el bloque de Comments (no añadir políticas como beneficios).
    for ul in soup.find_all("ul"):
        texts = [li.get_text(" ", strip=True) for li in ul.find_all("li")]
        if texts and any("discount" in t.lower() or "credit" in t.lower() or "upgrade" in t.lower() for t in texts):
            r.benefits = texts
            break
    # Marcamos valores que podrían requerir revisión antes de archivar.
    if not r.name: r.warnings.append("Falta el nombre del huésped")
    if not r.room: r.warnings.append("No se identificó la habitación")
    if r.total <= 0: r.warnings.append("Revisar el importe total extraído del correo")
    if not r.checkin or not r.checkout: r.warnings.append("Faltan fechas")
    if r.night_rates and not r.additional_services and abs(r.room_total - r.total) > Decimal("0.02"):
        r.warnings.append("La suma de tarifas por noche no coincide con el importe total")
    if r.nights:
        for room in r.rooms:
            if len(room.night_rates) != r.nights:
                r.warnings.append(f"La habitación {room.number} tiene {len(room.night_rates)} "
                                  f"noches desglosadas; se esperaban {r.nights}.")
    if len(r.rooms) > 4:
        r.warnings.append("Se detectaron más de cuatro habitaciones: revisar que el documento se imprima completo.")
    for room in r.rooms:
        if room.stated_total is not None and abs(room.stated_total - room.total) > Decimal("0.02"):
            r.warnings.append(f"El subtotal de la habitación {room.number} no concilia con Paraty.")
    return r


def fmt_usd(x: Decimal) -> str:
    return f"{x:,.2f} USD"


def set_cell_shading(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def para(p, text: str, bold: bool = False, color=None, size=9):
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    if color: run.font.color.rgb = RGBColor(*color)
    return run


def fmt_date(s: str) -> str:
    try: return datetime.fromisoformat(s).strftime("%d/%m/%Y")
    except ValueError: return s


def prepare_masked_image(image_bytes: bytes) -> bytes:
    """Normalize untrusted paste/upload into bounded PNG bytes.

    Security notice: this is NOT a PAN/CVV detection or masking function.
    The operator must redact sensitive information before pasting.
    """
    if len(image_bytes) > 6 * 1024 * 1024:
        raise ValueError("La imagen supera el límite de 6 MB.")
    try:
        with Image.open(BytesIO(image_bytes)) as loaded:
            im = ImageOps.exif_transpose(loaded)
            im.load()
            if im.width * im.height > 14_000_000:
                raise ValueError("La imagen es demasiado grande.")
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGB")
            im.thumbnail((1300, 1300))
            buf = BytesIO()
            im.save(buf, "PNG")
            return buf.getvalue()
    except (OSError, SyntaxError) as exc:
        raise ValueError("No se pudo leer la imagen. Usa PNG o JPG.") from exc


def generate_docx(r: Booking, card_type: str, tc: Decimal | None, last_four: str = "",
                  inssist_number: str = "", payment_image: bytes | None = None) -> bytes:
    d = Document()
    section = d.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    section.left_margin = section.right_margin = Inches(0.55)
    section.top_margin = Inches(0.53)
    section.bottom_margin = Inches(0.55)
    normal = d.styles["Normal"]
    normal.font.name, normal.font.size = "Arial", Pt(9)
    normal.paragraph_format.space_after = Pt(4)

    p = d.add_paragraph()
    para(p, "Booking reference: ", True, size=9)
    para(p, r.reference, size=9)
    if inssist_number.strip():
        para(p, "\nConfirmación INSSIST: ", True, color=(30, 75, 130), size=9)
        para(p, inssist_number.strip(), True, color=(30, 75, 130), size=9)
    para(p, "\nOrigin: " + r.origin + "\nDevice: " + r.device, size=8.5)

    header = d.add_table(rows=1, cols=2)
    header.autofit = False
    # Ancho útil Carta con márgenes de 0.55 in: 18.8 cm aprox.
    # La columna derecha mide 7.5 cm para alojar una foto de hasta 7 x 10 cm.
    header.columns[0].width = Cm(11.25)
    header.columns[1].width = Cm(7.5)
    left, right = header.rows[0].cells
    left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    left.width, right.width = Cm(11.25), Cm(7.5)
    # Quitar por completo el antiguo recuadro gris; conservar solo la imagen
    # sin fondo ni borde. Reducimos el relleno para que quepan 7 cm de ancho.
    tc_mar = right._tc.get_or_add_tcPr()
    margins = OxmlElement("w:tcMar")
    for edge in ("top", "left", "bottom", "right"):
        item = OxmlElement(f"w:{edge}")
        item.set(qn("w:w"), "30")
        item.set(qn("w:type"), "dxa")
        margins.append(item)
    tc_mar.append(margins)
    p = left.paragraphs[0]
    para(p, "Name: " + r.name, size=14)
    for label, value in [
        ("Email address", r.email), ("Telephone", r.phone),
        ("Address", r.location), ("Special requirements", r.requirements),
        ("Estimated Time of Arrival", r.arrival_time)
    ]:
        if value:
            q = left.add_paragraph()
            para(q, label + ": ", True, size=8)
            para(q, value, size=8)
    if payment_image:
        # Cuando la imagen vertical ocupa ~10 cm, aprovechamos el espacio
        # libre de la columna izquierda para no desperdiciar media página.
        q = left.add_paragraph()
        para(q, "Total price: ", True)
        para(q, fmt_usd(r.total))
        q = left.add_paragraph()
        para(q, "Taxes per stay: ", True)
        para(q, fmt_usd(r.taxes))
        q = left.add_paragraph()
        if r.deposit_required is not None:
            para(q, "Total pendiente según Paraty (NO es el depósito): ", True, size=9)
            para(q, fmt_usd(r.pending), size=9)
        else:
            para(q, "Pending payment: ", True, color=(200, 30, 30), size=11)
            para(q, fmt_usd(r.pending), True, color=(200, 30, 30), size=11)
        if r.deposit_required is not None:
            q = left.add_paragraph()
            para(q, f"Depósito requerido (promedio de {r.nights} noches + servicios): ",
                 True, color=(200, 30, 30), size=10)
            para(q, fmt_usd(r.deposit_required), True, color=(200, 30, 30), size=10)
            q = left.add_paragraph()
            para(q, "Alojamiento: " + fmt_usd(r.deposit_room_average) +
                 f" + {deposit_service_label(r)}: " + fmt_usd(r.deposit_services_charge), size=8)
        if card_type == "American Express" and tc is not None:
            mxn = (r.total * tc).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            q = left.add_paragraph()
            para(q, "MXN $" + f"{mxn:,.2f}" + "     tc " + format(tc, "f"), True,
                 color=(200, 30, 30), size=11)
            if r.deposit_required is not None:
                dep_mxn = (r.deposit_required * tc).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                q = left.add_paragraph()
                para(q, f"Depósito MXN ${dep_mxn:,.2f} (tc {tc})", True,
                     color=(30, 75, 130), size=9)
        q = left.add_paragraph()
        para(q, f"Number of Nights: {r.nights}\nCheck-In Date: {fmt_date(r.checkin)}\nCheck-Out Date: {fmt_date(r.checkout)}")
        q = left.add_paragraph()
        para(q, "Rooms", True, size=10)
        q = left.add_paragraph()
        for room in r.rooms:
            q = left.add_paragraph()
            para(q, f"Room {room.number}: {room.name} ({room.adults} Adults, "
                    f"{room.children} Children)", size=8)
        for label, val in [
            ("Room Selected", r.room), ("Selected Board Plan", r.meal_plan),
            ("Rate", r.rate), ("Promotion", r.promotion),
            ("Política detectada", policy_label(r)),
            ("Cancellation policy", r.cancellation_policy),
            ("Additional information of interest", r.requirements),
        ]:
            if val:
                q = left.add_paragraph()
                para(q, label + ": ", True, size=8)
                para(q, val, size=8)

    if payment_image:
        # Cuadro máximo 7 x 10 cm. Mantener proporciones; no estirar el recorte.
        p_img = right.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_img.paragraph_format.space_after = Pt(0)
        with Image.open(BytesIO(payment_image)) as img:
            ratio = min(7.0 / img.width, 10.0 / img.height)
            width_cm = img.width * ratio
            height_cm = img.height * ratio
        p_img.add_run().add_picture(BytesIO(payment_image),
                                    width=Cm(width_cm), height=Cm(height_cm))
    if card_type != "No identificado":
        q = right.add_paragraph()
        q.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        para(q, "Tipo: " + card_type, size=8)
    if last_four:
        q = right.add_paragraph()
        q.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        para(q, "Terminación: " + last_four, size=8)

    d.add_paragraph("_________________________________________________________________________________")
    if not payment_image:
        p = d.add_paragraph()
        para(p, "Total price: ", True)
        para(p, fmt_usd(r.total))
        p = d.add_paragraph()
        para(p, "Taxes per stay: ", True)
        para(p, fmt_usd(r.taxes))
        p = d.add_paragraph()
        if r.deposit_required is not None:
            para(p, "Total pendiente según Paraty (NO es el depósito): ", True, size=9)
            para(p, fmt_usd(r.pending), size=9)
        else:
            para(p, "Pending payment: ", True, color=(200, 30, 30), size=11)
            para(p, fmt_usd(r.pending), True, color=(200, 30, 30), size=11)
        if r.deposit_required is not None:
            p = d.add_paragraph()
            para(p, f"Depósito requerido (promedio de {r.nights} noches + servicios): ",
                 True, color=(200, 30, 30), size=10)
            para(p, fmt_usd(r.deposit_required), True, color=(200, 30, 30), size=10)
            p = d.add_paragraph()
            para(p, "Alojamiento: " + fmt_usd(r.deposit_room_average) +
                 f" + {deposit_service_label(r)}: " + fmt_usd(r.deposit_services_charge), size=8)
        if card_type == "American Express" and tc is not None:
            mxn = (r.total * tc).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            p = d.add_paragraph()
            para(p, "MXN $" + f"{mxn:,.2f}" + "     tc " + format(tc, "f"), True,
                 color=(200, 30, 30), size=11)
            if r.deposit_required is not None:
                dep_mxn = (r.deposit_required * tc).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                p = d.add_paragraph()
                para(p, f"Depósito MXN ${dep_mxn:,.2f} (tc {tc})", True,
                     color=(30, 75, 130), size=9)

    d.add_paragraph()
    if r.rooms:
        # Una tabla compacta por habitación conserva el desglose nocturno;
        # el depósito se calcula con el promedio del total de alojamiento.
        for room in r.rooms:
            p = d.add_paragraph()
            p.paragraph_format.space_after = Pt(1)
            para(p, f"Room {room.number}: {room.name} ({room.adults}-{room.children}-{room.infants})", True, size=9)
            tab = d.add_table(rows=1, cols=4)
            tab.style = "Table Grid"
            tab.alignment = WD_TABLE_ALIGNMENT.CENTER
            for cell, title in zip(tab.rows[0].cells, ["Fecha", "Tarifa", "Descuento", "Total noche"]):
                set_cell_shading(cell, "2261A8")
                para(cell.paragraphs[0], title, True, color=(255,255,255), size=8)
            for night in room.night_rates:
                cells = tab.add_row().cells
                for cell, content in zip(cells, [night.date, fmt_usd(night.original),
                                                "-" + fmt_usd(night.discount), fmt_usd(night.net)]):
                    para(cell.paragraphs[0], content, size=8)
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.space_after = Pt(1)
            para(p, f"Subtotal habitación {room.number}: " + fmt_usd(room.total), True, size=8)
        if len(r.rooms) > 1:
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            para(p, "Subtotal todas las habitaciones: " + fmt_usd(r.room_total), True, size=9)
        if r.deposit_required is not None:
            p = d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            para(p, "DEPÓSITO FLEXIBLE: " + fmt_usd(r.deposit_required) +
                 f"  (alojamiento / {r.nights} noches + {deposit_service_label(r)})", True,
                 color=(30, 75, 130), size=9)
            q = d.add_paragraph()
            q.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            para(q, "Promedio alojamiento: " + fmt_usd(r.deposit_room_average) +
                 f" | {deposit_service_label(r)} al primer cargo: " + fmt_usd(r.deposit_services_charge),
                 size=8)

    if r.additional_services:
        p = d.add_paragraph()
        para(p, "ADDITIONAL SERVICES / SERVICIOS ADICIONALES", True, size=10)
        tab_services = d.add_table(rows=1, cols=4)
        tab_services.style = "Table Grid"
        tab_services.alignment = WD_TABLE_ALIGNMENT.CENTER
        for cell, title in zip(tab_services.rows[0].cells, ["Servicio", "Cantidad", "Fecha", "Precio"]):
            set_cell_shading(cell, "2261A8")
            para(cell.paragraphs[0], title, True, color=(255, 255, 255), size=8)
        for item in r.additional_services:
            cells = tab_services.add_row().cells
            for cell, content in zip(cells, [item.name, str(item.quantity), item.date, f"{item.price:,.2f} {item.currency}"]):
                para(cell.paragraphs[0], content, size=8)
        p = d.add_paragraph()
        para(p, "Subtotal servicios (antes de IVA): ", True, size=8)
        para(p, fmt_usd(r.services_total), size=8)
        p = d.add_paragraph()
        para(p, f"Impuesto {tax_percent_label(r)} de servicios: ", True, size=8)
        para(p, fmt_usd(r.services_tax), size=8)
        p = d.add_paragraph()
        para(p, "Total servicios con IVA (al primer cargo): ", True, size=8)
        para(p, fmt_usd(r.services_with_tax), size=8)
        if abs(r.extra_difference) > Decimal("0.02"):
            p = d.add_paragraph()
            para(p, "Diferencia con el total Paraty (revisar, NO se añade al depósito): ",
                 True, color=(160, 75, 0), size=8)
            para(p, fmt_usd(r.extra_difference), color=(160, 75, 0), size=8)
        p = d.add_paragraph()
        para(p, "Total oficial Paraty: ", True)
        para(p, fmt_usd(r.total), True)

    if not payment_image:
        p = d.add_paragraph()
        para(p, f"Number of Nights: {r.nights}\nCheck-In Date: {fmt_date(r.checkin)}\nCheck-Out Date: {fmt_date(r.checkout)}")
    if not payment_image:
        p = d.add_paragraph()
        para(p, "Rooms", True, size=10)
        for room in r.rooms:
            d.add_paragraph(f"Room {room.number}: {room.name} "
                            f"({room.adults} Adults, {room.children} Children)")
        for label, val in [
            ("Room Selected", r.room), ("Selected Board Plan", r.meal_plan),
            ("Rate", r.rate), ("Promotion", r.promotion),
            ("Política detectada", policy_label(r)),
            ("Cancellation policy", r.cancellation_policy),
            ("Additional information of interest", r.requirements),
        ]:
            if val:
                p=d.add_paragraph()
                para(p, label + ": ", True)
                para(p, val)
    if r.comments:
        p=d.add_paragraph()
        para(p, "Comments: ", True)
        para(p, r.comments)
    if r.benefits:
        p = d.add_paragraph()
        para(p, "BENEFITS", True)
        for benefit in r.benefits:
            p = d.add_paragraph(style="List Bullet")
            para(p, benefit, size=8)
    if r.terms:
        p=d.add_paragraph()
        para(p,"BOOKING TERMS AND CONDITIONS",True)
        p=d.add_paragraph()
        para(p,r.terms,size=8)
    p=d.add_paragraph()
    if payment_image:
        para(p,"Imagen de referencia enmascarada. Verificación bancaria exclusivamente en Paraty Manager.",False,(90,90,90),8)
    else:
        para(p,"Tarjeta no incluida: consultar la plataforma de pago autorizada del hotel.",False,(90,90,90),8)
    data=BytesIO()
    d.save(data)
    return data.getvalue()


def generate_print_html(r: Booking, card_type: str, tc: Decimal | None,
                        last_four: str = "", inssist_number: str = "",
                        payment_image: bytes | None = None) -> str:
    """Vista imprimible desde el navegador, sin descargar Word ni enviar datos.

    Al pulsar el botón del visor se abre el diálogo de impresión del navegador.
    Se imprime solamente este HTML de la reservación, no la interfaz Streamlit.
    """
    h = html_escape
    def p_text(s: object) -> str:
        return h(str(s), quote=True)
    def pair(label: str, value: str) -> str:
        return (f'<p><b>{h(label)}:</b> {p_text(value)}</p>') if value else ""
    img_html = ""
    if payment_image:
        encoded = base64.b64encode(payment_image).decode("ascii")
        img_html = (f'<img alt="Imagen de referencia enmascarada" '
                    f'src="data:image/png;base64,{encoded}" '
                    f'style="display:block;max-width:7cm;max-height:10cm;width:auto;height:auto;margin-left:auto" />')
    payment_label = ""
    if card_type != "No identificado":
        payment_label += pair("Tipo", card_type)
    if last_four:
        payment_label += pair("Terminación", last_four)
    image_column = f'<div class="payment-image">{img_html}{payment_label}</div>'

    room_rates_html = ""
    if r.rooms:
        blocks = []
        for room in r.rooms:
            rows = "".join(
                f'<tr><td>{p_text(x.date)}</td><td>{p_text(fmt_usd(x.original))}</td>'
                f'<td>-{p_text(fmt_usd(x.discount))}</td><td>{p_text(fmt_usd(x.net))}</td></tr>'
                for x in room.night_rates
            )
            blocks.append(
                f'<h3>Room {room.number}: {p_text(room.name)} '
                f'({room.adults}-{room.children}-{room.infants})</h3>'
                '<table><thead><tr><th>Fecha</th><th>Tarifa</th><th>Descuento</th>'
                f'<th>Total noche</th></tr></thead><tbody>{rows}</tbody></table>'
                f'<p class="right"><b>Subtotal habitación {room.number}: '
                f'{p_text(fmt_usd(room.total))}</b></p>'
            )
        if len(r.rooms) > 1:
            blocks.append(f'<p class="right"><b>Subtotal habitaciones: '
                          f'{p_text(fmt_usd(r.room_total))}</b></p>')
        room_rates_html = "".join(blocks)

    services_html = ""
    if r.additional_services:
        rows = "".join(
            f'<tr><td>{p_text(x.name)}</td><td>{x.quantity}</td><td>{p_text(x.date)}</td>'
            f'<td>{p_text(f"{x.price:,.2f} {x.currency}")}</td></tr>'
            for x in r.additional_services
        )
        diff_html = ""
        if abs(r.extra_difference) > Decimal("0.02"):
            diff_html = ('<p class="note"><b>Diferencia con el total Paraty (revisar, NO se añade al depósito):</b> '
                         f'{p_text(fmt_usd(r.extra_difference))}</p>')
        services_html = f"""<h3>ADDITIONAL SERVICES / SERVICIOS ADICIONALES</h3>
        <table><thead><tr><th>Servicio</th><th>Cantidad</th><th>Fecha</th><th>Precio</th></tr></thead>
        <tbody>{rows}</tbody></table>
        <p><b>Subtotal servicios (sin IVA):</b> {p_text(fmt_usd(r.services_total))}</p>
        <p><b>Impuesto {tax_percent_label(r)} servicios:</b> {p_text(fmt_usd(r.services_tax))}</p>
        <p><b>Total servicios con IVA (al primer cargo):</b> {p_text(fmt_usd(r.services_with_tax))}</p>
        {diff_html}
        <p><b>Total oficial Paraty: {p_text(fmt_usd(r.total))}</b></p>"""

    benefits_html = ""
    if r.benefits:
        benefits_html = "<h3>BENEFITS</h3><ul>" + "".join(
            f"<li>{p_text(item)}</li>" for item in r.benefits) + "</ul>"
    tc_html = ""
    if card_type == "American Express" and tc is not None:
        mxn = (r.total * tc).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        tc_html = f'<p class="red strong">MXN ${mxn:,.2f} &nbsp;&nbsp; tc {p_text(format(tc,"f"))}</p>'
        if r.deposit_required is not None:
            dep_mxn = (r.deposit_required * tc).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            tc_html += f'<p class="deposit">Depósito MXN ${dep_mxn:,.2f} (tc {p_text(tc)})</p>'

    amount_html = f'''<p><b>Total price:</b> {p_text(fmt_usd(r.total))}</p>
     <p><b>Taxes per stay:</b> {p_text(fmt_usd(r.taxes))}</p>
     {f'<p><b>Total pendiente según Paraty (NO es el depósito):</b> {p_text(fmt_usd(r.pending))}</p>' if r.deposit_required is not None else f'<p class="red strong">Pending payment: {p_text(fmt_usd(r.pending))}</p>'}
     {f'<p class="red strong"><b>DEPÓSITO FLEXIBLE (promedio alojamiento + {deposit_service_label(r)}): {p_text(fmt_usd(r.deposit_required))}</b></p><p class="small">Alojamiento: {p_text(fmt_usd(r.deposit_room_average))} + {deposit_service_label(r)}: {p_text(fmt_usd(r.deposit_services_charge))}</p>' if r.deposit_required is not None else ''}
     {tc_html}'''
    nights_html = f'''<p>Number of Nights: {r.nights}</p>
     <p>Check-In Date: {p_text(fmt_date(r.checkin))}</p>
     <p>Check-Out Date: {p_text(fmt_date(r.checkout))}</p>'''
    room_meta_html = '<h3>Rooms</h3>' + ''.join(
        f'<p>Room {room.number}: {p_text(room.name)} '
        f'({room.adults} Adults, {room.children} Children)</p>' for room in r.rooms
    )
    room_meta_html += (pair("Room Selected",r.room) + pair("Selected Board Plan",r.meal_plan)
                       + pair("Rate",r.rate) + pair("Promotion",r.promotion)
                       + pair("Política detectada",policy_label(r))
                       + pair("Cancellation policy",r.cancellation_policy)
                       + pair("Additional information of interest",r.requirements))
    in_top = (amount_html + nights_html + room_meta_html) if payment_image else ""
    after_top = "" if payment_image else amount_html
    after_rates = "" if payment_image else (nights_html + room_meta_html)

    return f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8" />
<style>
@page {{ size: letter; margin: .53in .55in .55in .55in; }}
* {{ box-sizing: border-box; }}
html, body {{ margin:0; padding:0; font-family: Arial, sans-serif; color:#111; }}
body {{ font-size: 9pt; }}
.toolbar {{ padding:9px 12px; background:#fff; text-align:left; }}
button {{ cursor:pointer; background:#1452a3; color:white; border:0; border-radius:6px; padding:9px 16px; font-size:14px; font-weight:600; }}
button:hover {{ background:#0d3f83; }}
.instructions {{ font-size:12px; margin-left:10px; color:#596273; }}
.paper {{ max-width: 21.6cm; margin:0 auto; padding: .53in .55in .55in .55in; }}
.top {{ display:grid; grid-template-columns:minmax(0,1fr) 7.5cm; column-gap:.15cm; align-items:start; }}
.payment-image {{ text-align:right; padding:0; margin:0; background:transparent; border:0; }}
h2 {{ font-size:14pt; font-weight:normal; margin:5px 0; }}
h3 {{ font-size:10pt; margin:12px 0 5px; break-after:avoid; }}
p {{ margin:4px 0; }}
.red {{ color:#c91c24; }}
.deposit {{ color:#1e4b82; font-weight:bold; font-size:10pt; }}
.strong {{ font-weight:700; font-size:11pt; }}
.rule {{ border:0; border-top:1px solid #444; margin:10px 0; }}
.right {{ text-align:right; }}
.note {{ color:#a04b00; font-size:8pt; }}
.small {{ font-size:7.5pt; }}
table {{ width:100%; border-collapse:collapse; table-layout:fixed; margin:3px 0; font-size:8pt; }}
th,td {{ border:1px solid #4c4c4c; padding:3px 4px; overflow-wrap:anywhere; }}
th {{ background:#2261a8; color:#fff; text-align:left; }}
tr {{ break-inside:avoid; }}
ul {{ margin:3px 0 10px; padding-left:20px; font-size:8pt; }}
.terms {{ font-size:8pt; }}
@media print {{
  .toolbar {{ display:none !important; }}
  .paper {{ max-width:none; margin:0; padding:0; }}
  body {{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }}
}}
@media screen {{
  .paper {{ display:none; }}
}}
</style></head><body>
<div class="toolbar"><button type="button" onclick="window.print()">🖨️ Imprimir confirmación</button>
<span class="instructions">Abre el diálogo de impresión de tu navegador.</span></div>
<article class="paper">
 <p><b>Booking reference:</b> {p_text(r.reference)}</p>
 {pair("Confirmación INSSIST", inssist_number)}
 <p>Origin: {p_text(r.origin)}</p><p>Device: {p_text(r.device)}</p>
 <div class="top"><div><h2>Name: {p_text(r.name)}</h2>
 {pair("Email address",r.email)}{pair("Telephone",r.phone)}
 {pair("Address",r.location)}{pair("Special requirements",r.requirements)}
 {pair("Estimated Time of Arrival",r.arrival_time)}
 {in_top}
 </div>{image_column}</div>
 <hr class="rule" />
 {after_top}
 {room_rates_html}
 {services_html}
 {after_rates}
 {pair("Comments",r.comments)}
 {benefits_html}
 <h3>BOOKING TERMS AND CONDITIONS</h3><p class="terms">{p_text(r.terms)}</p>
 </article></body></html>"""


# OAuth de escritorio: apropiado únicamente cuando la app Streamlit corre en la
# misma computadora que abre Gmail. Para hosting multiusuario utilizar OAuth web,
# tokens por usuario y almacenamiento seguro (no reutilizar token.json común).
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
APP_DIR = Path(__file__).resolve().parent
CREDENTIALS_PATH = APP_DIR / "credentials.json"
TOKEN_PATH = APP_DIR / "token.json"


def load_gmail_service(allow_login: bool = False):
    """Obtiene Gmail API, con permiso de solo lectura. Nunca guarda correos."""
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    creds = None
    if TOKEN_PATH.is_file():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), GMAIL_SCOPES)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    if not creds or not creds.valid:
        if not allow_login:
            return None
        if not CREDENTIALS_PATH.is_file():
            raise FileNotFoundError("Falta credentials.json de Google Cloud en la carpeta de la aplicación.")
        flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), GMAIL_SCOPES)
        creds = flow.run_local_server(port=0, open_browser=True)
    # Token solo en la computadora local; proteger y nunca subir a GitHub.
    TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
    if os.name != "nt":
        TOKEN_PATH.chmod(0o600)
    return build("gmail", "v1", credentials=creds, cache_discovery=False)


def gmail_list_booking_emails(service, query: str, limit: int = 25) -> list[dict]:
    """Devuelve metadatos; el cuerpo completo solo se lee al seleccionar uno."""
    response = service.users().messages().list(
        userId="me", q=query, maxResults=limit
    ).execute()
    items = []
    for item in response.get("messages", []):
        data = service.users().messages().get(
            userId="me", id=item["id"], format="metadata",
            metadataHeaders=["Subject", "Date", "From"]
        ).execute()
        headers = {h["name"].casefold(): h["value"] for h in data.get("payload", {}).get("headers", [])}
        subject = str(make_header(decode_header(headers.get("subject", "(sin asunto)"))))
        # La búsqueda Gmail ya filtra; este control evita resultados ajenos.
        if "booking confirmation" not in subject.casefold():
            continue
        items.append({"id": item["id"], "subject": subject,
                      "date": headers.get("date", ""), "from": headers.get("from", "")})
    return items


def gmail_read_eml(service, gmail_message_id: str) -> bytes:
    msg = service.users().messages().get(
        userId="me", id=gmail_message_id, format="raw"
    ).execute()
    raw = msg.get("raw")
    if not raw:
        raise ValueError("Gmail no devolvió contenido para este mensaje.")
    return base64.urlsafe_b64decode(raw + "=" * (-len(raw) % 4))


def main():
    import streamlit as st
    from streamlit_paste_button import paste_image_button

    st.set_page_config(page_title="Confirmaciones Paraty", layout="wide")
    st.title("Confirmaciones Paraty → Word")
    st.caption("Gmail de solo lectura o archivo .eml. Se genera un Word editable, sin guardar la reservación en una base de datos.")

    with st.sidebar:
        st.header("Origen de confirmación")
        source_mode = st.radio("Selecciona el origen", ["Gmail (directo)", "Archivo .eml"], key="source_mode")
        st.caption("Gmail usa OAuth local. No se requiere descargar el .eml.")

    mail_data = None
    source_key = None
    if source_mode == "Archivo .eml":
        eml = st.file_uploader("Adjunta un correo de confirmación Paraty (.eml)", type=["eml"])
        if eml is not None:
            mail_data = eml.getvalue()
            # hash para detectar que el operador cambió el correo
            import hashlib
            source_key = "eml-" + hashlib.sha256(mail_data).hexdigest()[:16]
        else:
            st.info("Carga el correo .eml o selecciona Gmail en la barra lateral.")
            st.stop()
    else:
        if not CREDENTIALS_PATH.is_file() and not TOKEN_PATH.is_file():
            st.warning("Primero configura Gmail OAuth: coloca `credentials.json` en la carpeta de la aplicación. Revisa la guía adjunta.")
        service = None
        try:
            service = load_gmail_service(allow_login=False)
        except Exception as exc:
            st.error(f"No fue posible leer la sesión de Gmail: {exc}")
        if service is None:
            if st.button("Conectar Gmail (solo lectura)"):
                try:
                    service = load_gmail_service(allow_login=True)
                    st.success("Gmail conectado.")
                except Exception as exc:
                    st.error(f"No se pudo conectar Gmail: {exc}")
            if service is None:
                st.info("Si aún no configuraste Gmail, usa la opción Archivo .eml para continuar.")
                st.stop()

        query = st.text_input(
            "Buscar en Gmail", value='subject:"Booking confirmation" newer_than:90d',
            help="Puedes añadir un número de reservación o cambiar el rango de fechas usando la sintaxis de búsqueda de Gmail."
        )
        refresh = st.button("Buscar / actualizar confirmaciones")
        if refresh or st.session_state.get("gmail_query") != query:
            try:
                st.session_state["gmail_results"] = gmail_list_booking_emails(service, query)
                st.session_state["gmail_query"] = query
            except Exception as exc:
                st.error(f"No se pudo consultar la bandeja: {exc}")
                st.stop()
        results = st.session_state.get("gmail_results", [])
        if not results:
            st.info("No hay mensajes en la búsqueda. Puedes cambiar el filtro o cargar un .eml.")
            st.stop()
        selected_id = st.selectbox(
            "Selecciona el correo Paraty", [x["id"] for x in results],
            format_func=lambda k: next((f'{x["subject"]}  |  {x["date"][:25]}' for x in results if x["id"] == k), k)
        )
        source_key = "gmail-" + selected_id
        # No mezclar correos de reservas distintas dentro de una sesión.
        if st.session_state.get("loaded_mail_source") != source_key:
            try:
                mail_data = gmail_read_eml(service, selected_id)
                st.session_state["loaded_mail_data"] = mail_data
                st.session_state["loaded_mail_source"] = source_key
            except Exception as exc:
                st.error(f"No se pudo leer el correo seleccionado: {exc}")
                st.stop()
        else:
            mail_data = st.session_state.get("loaded_mail_data")

    if st.session_state.get("active_booking_source") != source_key:
        st.session_state["active_booking_source"] = source_key
        st.session_state.pop("payment_image", None)
        st.session_state.pop("last_image_upload_hash", None)
        st.session_state["image_upload_version"] = 0
        st.session_state["paste_key_version"] = 0
        st.session_state["insist_number"] = ""  # campo opcional por reservación
        st.session_state["card_type"] = "No identificado"
        st.session_state["last_four"] = ""
        st.session_state["service_tax_choice"] = "Selecciona impuesto..."
        st.session_state.pop("custom_service_tax", None)
    try:
        booking = parse_paraty_eml(mail_data)
    except Exception as exc:
        st.error(f"No se pudo interpretar este correo como confirmación Paraty: {exc}")
        st.stop()

    # Preguntar impuesto SOLO cuando el correo trae Additional Services.
    # No inferir automáticamente 16%/30% porque los correos varían.
    if booking.additional_services:
        st.subheader("Impuestos de servicios adicionales")
        tax_choice = st.selectbox(
            "¿Qué porcentaje de impuesto se aplica a los servicios de esta reservación?",
            ["Selecciona impuesto...", "16%", "30%", "Otro porcentaje"],
            key="service_tax_choice",
            help="Este impuesto se aplica a los servicios, NO a las habitaciones. "
                 "Debes seleccionar uno antes de descargar o imprimir."
        )
        chosen_tax_rate = None
        if tax_choice == "16%":
            chosen_tax_rate = Decimal("0.16")
        elif tax_choice == "30%":
            chosen_tax_rate = Decimal("0.30")
        elif tax_choice == "Otro porcentaje":
            custom_tax = st.number_input(
                "Impuesto personalizado (%)", min_value=0.0, max_value=100.0,
                value=16.0, step=0.5, format="%.2f", key="custom_service_tax"
            )
            chosen_tax_rate = Decimal(str(custom_tax)) / Decimal("100")
        recalculate_booking_finances(booking, chosen_tax_rate)
        if chosen_tax_rate is not None:
            st.info(f"Impuesto elegido para servicios: {tax_percent_label(booking)}. "
                    "Se aplicará al depósito y al Word.")

    if booking.warnings:
        for warning in booking.warnings:
            st.warning(warning)
    st.subheader(f"Paraty: {booking.reference}")
    st.write("**Detección basada en políticas de cancelación y pago:**")
    if booking.policy_status == "flexible":
        st.success(policy_label(booking) + " — " + booking.policy_reason)
    elif booking.policy_status == "no_reembolsable":
        st.info(policy_label(booking) + " — " + booking.policy_reason)
    else:
        st.error(policy_label(booking) + " — " + booking.policy_reason)
    st.caption("Nombre de tarifa Paraty: " + (booking.rate or "No identificado") +
               ". El nombre no determina el anticipo; prevalecen las políticas del correo.")
    col1, col2, col3 = st.columns(3)
    col1.metric("Importe", fmt_usd(booking.total))
    col2.metric("Noches", booking.nights)
    col3.metric("Habitaciones", len(booking.rooms))
    st.write("**Huésped:**", booking.name)
    st.write("**Estancia:**", fmt_date(booking.checkin), "—", fmt_date(booking.checkout))
    if booking.rooms:
        for room in booking.rooms:
            st.markdown(f"**Habitación {room.number}: {room.name}** — "
                        f"{room.adults} adultos, {room.children} menores")
            st.dataframe([{"Fecha": n.date, "Tarifa USD": float(n.original),
                          "Descuento USD": float(n.discount), "Neto USD": float(n.net)}
                         for n in room.night_rates], hide_index=True, use_container_width=True)
            st.caption(f"Subtotal habitación {room.number}: {fmt_usd(room.total)}")
        st.write("**Subtotal todas las habitaciones:**", fmt_usd(booking.room_total))
    if booking.deposit_required is not None:
        st.success(f"DEPÓSITO FLEXIBLE — promedio de una noche de la reserva "
                   f"({len(booking.rooms)} habitación(es)) + {deposit_service_label(booking)}: "
                   f"{fmt_usd(booking.deposit_required)}")
        st.caption(f"Alojamiento a prorratear: {fmt_usd(booking.deposit_base_amount)} "
                   f"÷ {booking.nights} noches = {fmt_usd(booking.deposit_room_average)}; "
                   f"{deposit_service_label(booking)} al primer cargo: "
                   f"{fmt_usd(booking.deposit_services_charge)}. "
                   "El total de Paraty ya incluye esos suplementos, por eso se "
                   "separan antes de prorratear para no cobrarlos dos veces.")
    if booking.additional_services:
        st.subheader("Additional Services / Servicios adicionales")
        if booking.services_tax_rate is not None:
            st.dataframe([{"Servicio": x.name, "Cantidad": x.quantity, "Fecha": x.date,
                           "Base": f"{x.price:,.2f} {x.currency}",
                           f"Impuesto {tax_percent_label(booking)}":
                               f"{(x.price * booking.services_tax_rate).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP):,.2f} {x.currency}",
                           "Total con impuesto":
                               f"{(x.price + (x.price * booking.services_tax_rate).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)):,.2f} {x.currency}"}
                          for x in booking.additional_services], hide_index=True, use_container_width=True)
            st.write("**Subtotal servicios (sin impuestos):**", fmt_usd(booking.services_total))
            st.write(f"**Impuesto sobre servicios ({tax_percent_label(booking)}):**", fmt_usd(booking.services_tax))
            st.write("**Servicios con impuesto al depósito:**", fmt_usd(booking.services_with_tax))
            if abs(booking.extra_difference) > Decimal("0.02"):
                st.warning(f"El total Paraty difiere del desglose con impuesto "
                           f"{tax_percent_label(booking)} por {fmt_usd(booking.extra_difference)}. "
                           "Revisar el importe; la diferencia no se suma automáticamente al depósito.")
        else:
            st.info("Selecciona el impuesto para ver el importe total de cada servicio.")
        st.write("**Total oficial de la reservación:**", fmt_usd(booking.total))
    st.divider()

    inssist_number = st.text_input(
        "Número de confirmación INSSIST (opcional)",
        key="insist_number", max_chars=60,
        help="Se imprimirá junto a la referencia de Paraty, en la parte superior del Word."
    ).strip()
    card_type = st.selectbox("Tipo de tarjeta (verificar en Paraty Manager)",
                             ["No identificado", "American Express", "Otra tarjeta"],
                             key="card_type")
    tc = None
    if card_type == "American Express":
        # Tipo de cambio fijo por configuración. Editar DEFAULT_EXCHANGE_RATE
        # al principio de este archivo para modificarlo para toda la aplicación.
        tc = DEFAULT_EXCHANGE_RATE
        mxn = (booking.total * tc).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        st.info(f"Tipo de cambio fijo configurado: {tc} MXN por USD "
                "(editable en DEFAULT_EXCHANGE_RATE dentro de app.py).")
        st.success(f"Equivalente: MXN ${mxn:,.2f}")
    last_four = st.text_input("Últimos cuatro dígitos (opcional)",
                              max_chars=4, key="last_four")
    bad_four = bool(last_four and (not last_four.isdigit() or len(last_four) != 4))
    if bad_four:
        st.error("Escribe exactamente cuatro dígitos o deja el campo vacío.")

    st.subheader("Imagen de referencia (esquina superior derecha, 7 × 10 cm)")
    st.warning("Por seguridad, NO pegues capturas que muestren número completo de tarjeta, CVV, fecha de vencimiento u otros datos bancarios sensibles. Ocúltalos antes de copiar la imagen. El Word no es una bóveda PCI.")
    st.caption("Copia una imagen previamente enmascarada (Windows + Shift + S) y pulsa Pegar imagen. También puedes subir un PNG/JPG enmascarado.")
    p = paste_image_button("Pegar imagen copiada del portapapeles",
                           key=f"paste_image_{source_key}_{st.session_state.get('paste_key_version', 0)}",
                           errors="raise")
    if p.image_data is not None:
        buf = BytesIO()
        p.image_data.save(buf, format="PNG")
        try:
            st.session_state["payment_image"] = prepare_masked_image(buf.getvalue())
        except ValueError as exc:
            st.error(str(exc))
    # No mostrar permanentemente el cuadro grande de drag-and-drop: la captura
    # del portapapeles es la vía principal. La carga queda disponible como respaldo.
    with st.expander("Alternativa: subir imagen PNG/JPG", expanded=False):
        uploaded_image = st.file_uploader("Seleccionar imagen ya enmascarada", type=["png", "jpg", "jpeg"], key=f"image_upload_{source_key}_{st.session_state.get('image_upload_version', 0)}")
    if uploaded_image is not None:
        try:
            import hashlib
            image_bytes = uploaded_image.getvalue()
            upload_hash = hashlib.sha256(image_bytes).hexdigest()
            if st.session_state.get("last_image_upload_hash") != (source_key, upload_hash):
                st.session_state["payment_image"] = prepare_masked_image(image_bytes)
                st.session_state["last_image_upload_hash"] = (source_key, upload_hash)
        except ValueError as exc:
            st.error(str(exc))
    payment_image = st.session_state.get("payment_image")
    if payment_image:
        st.image(payment_image, caption="Imagen para Word: espacio máximo 7 cm ancho × 10 cm alto, sin recuadro gris", width=360)
        if st.button("Quitar imagen"):
            st.session_state.pop("payment_image", None)
            st.session_state.pop("last_image_upload_hash", None)
            st.session_state["image_upload_version"] = st.session_state.get("image_upload_version", 0) + 1
            st.session_state["paste_key_version"] = st.session_state.get("paste_key_version", 0) + 1
            st.rerun()
    acknowledge = st.checkbox(
        "Confirmo que la imagen NO contiene número completo, CVV ni fecha de vencimiento",
        value=False,
        help="No es una verificación automática; el personal debe revisar el contenido antes de descargar."
    ) if payment_image else True
    st.caption("Los datos de tarjeta completos se consultan exclusivamente en Paraty Manager; Gmail no los incluye.")

    disabled = ((card_type == "American Express" and tc is None) or bad_four or
                (bool(payment_image) and not acknowledge) or
                (booking.flexible_rate and booking.deposit_required is None) or
                (bool(booking.additional_services) and booking.services_tax_rate is None))
    if not disabled:
        try:
            output = generate_docx(booking, card_type, tc, last_four,
                                   inssist_number=inssist_number, payment_image=payment_image)
            col_word, col_print = st.columns([1, 1], gap="medium")
            with col_word:
                st.download_button("Descargar confirmación Word", output,
                                   file_name=f"Confirmacion_{booking.reference}.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   type="primary", use_container_width=True)
            with col_print:
                st.caption("Impresión directa: haz clic en el botón azul que aparece debajo.")
            # Streamlit se ejecuta en un navegador: no puede mandar un .docx a la
            # impresora de una computadora remota. Imprimimos el mismo contenido
            # desde una vista HTML local que abre el diálogo del navegador.
            import streamlit.components.v1 as components
            print_html = generate_print_html(booking, card_type, tc, last_four,
                                             inssist_number=inssist_number,
                                             payment_image=payment_image)
            components.html(print_html, height=66, scrolling=False)
        except Exception as exc:
            st.error(f"No se pudo generar el Word: {exc}")
    else:
        st.info("Completa los datos requeridos antes de descargar el documento.")


if __name__ == "__main__":
    main()
