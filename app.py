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
from PIL import Image, ImageOps
import re

from bs4 import BeautifulSoup
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


@dataclass
class Night:
    date: str
    original: Decimal
    discount: Decimal
    net: Decimal


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
    promotion: str = ""
    comments: str = ""
    benefits: list[str] = field(default_factory=list)
    terms: str = ""
    night_rates: list[Night] = field(default_factory=list)
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
    r.location = field_value(soup, "Address").lstrip(": ").strip()
    r.total = money(field_value(soup, "Total price"))
    # En Paraty «Pending payment» es texto del <span>, no una etiqueta <b>.
    pending_el = soup.find(string=lambda z: z and "Pending payment:" in z)
    if pending_el:
        r.pending = money(pending_el.parent.get_text(" ", strip=True).split(":", 1)[-1])
    tax_b = soup.find(["b", "strong"], string=lambda t: t and "Taxes per stay" in t)
    if tax_b:
        r.taxes = money(tax_b.parent.get_text(" ", strip=True).split(":", 1)[-1])
    for attr, item in [("checkin", "checkinDate"), ("checkout", "checkoutDate")]:
        el = soup.find(attrs={"itemprop": item})
        if el and el.get("content"):
            setattr(r, attr, el.get("content")[:10])
    if r.checkin and r.checkout:
        r.nights = (datetime.fromisoformat(r.checkout) - datetime.fromisoformat(r.checkin)).days
    r.adults = int((re.search(r"(\d+)\s+Adults", soup.get_text(" ", strip=True)) or [None, 0])[1])
    r.children = int((re.search(r"(\d+)\s+Children", soup.get_text(" ", strip=True)) or [None, 0])[1])
    # En el correo Paraty la tabla de tarifas tiene una subtabla por fecha.
    old = soup.select_one(".old_price_per_day_table > table")
    if old:
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
                    r.night_rates.append(Night(
                        date_str, money(amounts[0]), abs(money(amounts[1])), money(amounts[2])
                    ))
    # Beneficios: lista bajo el bloque de Comments (no añadir políticas como beneficios).
    for ul in soup.find_all("ul"):
        texts = [li.get_text(" ", strip=True) for li in ul.find_all("li")]
        if texts and any("discount" in t.lower() or "credit" in t.lower() or "upgrade" in t.lower() for t in texts):
            r.benefits = texts
            break
    # Conserva únicamente términos textuales suministrados en el correo.
    terms_b = soup.find(["b", "strong"], string=lambda t: t and "BOOKING TERMS AND CONDITIONS" in t)
    if terms_b:
        parts = []
        for sib in terms_b.next_siblings:
            if getattr(sib, "name", None) in ["div", "p"]:
                parts.append(sib.get_text(" ", strip=True))
            elif isinstance(sib, str):
                parts.append(sib.strip())
        r.terms = " ".join(x for x in parts if x)
        r.terms = re.sub(r"\s+", " ", r.terms)
    # Marcamos valores que podrían requerir revisión antes de archivar.
    if not r.name: r.warnings.append("Falta el nombre del huésped")
    if not r.room: r.warnings.append("No se identificó la habitación")
    if r.total <= 0: r.warnings.append("Revisar el importe total extraído del correo")
    if not r.checkin or not r.checkout: r.warnings.append("Faltan fechas")
    if r.night_rates and sum((x.net for x in r.night_rates), Decimal("0")) != r.total:
        r.warnings.append("La suma de tarifas por noche no coincide con el importe total")
    if r.nights and r.night_rates and len(r.night_rates) != r.nights:
        r.warnings.append("El número de desglose de noches no coincide con la estancia")
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
    header.columns[0].width = Inches(5.1)
    header.columns[1].width = Inches(2.1)
    left, right = header.rows[0].cells
    left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    for idx, cell in enumerate((left, right)):
        cell.width = Inches(5.1 if idx == 0 else 2.1)
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
    set_cell_shading(right, "F1F4F8")
    right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    para(right.paragraphs[0], "REFERENCIA DE PAGO", True, size=8)
    if payment_image:
        p_img = right.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        # Conservar el encabezado en la esquina superior derecha.
        with Image.open(BytesIO(payment_image)) as img:
            scale = min(1.85 / img.width, 1.30 / img.height)
            display_width = max(img.width * scale, 0.1)
            display_height = max(img.height * scale, 0.1)
        p_img.add_run().add_picture(BytesIO(payment_image),
                                    width=Inches(display_width), height=Inches(display_height))
        p_img.paragraph_format.space_after = Pt(2)
        para(right.add_paragraph(), "Imagen enmascarada", size=7.5)
    else:
        para(right.add_paragraph(), "Consultar tarjeta en Paraty Manager", size=8)
    if card_type != "No identificado":
        para(right.add_paragraph(), "Tipo: " + card_type, size=8)
    if last_four:
        para(right.add_paragraph(), "Terminación: " + last_four, size=8)

    d.add_paragraph("_________________________________________________________________________________")
    p = d.add_paragraph()
    para(p, "Total price: ", True)
    para(p, fmt_usd(r.total))
    p = d.add_paragraph()
    para(p, "Taxes per stay: ", True)
    para(p, fmt_usd(r.taxes))
    p = d.add_paragraph()
    para(p, "Pending payment: ", True, color=(200, 30, 30), size=11)
    para(p, fmt_usd(r.pending), True, color=(200, 30, 30), size=11)
    if card_type == "American Express" and tc is not None:
        mxn = (r.total * tc).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        p = d.add_paragraph()
        para(p, "MXN $" + f"{mxn:,.2f}" + "     tc " + format(tc, "f"), True, color=(200, 30, 30), size=11)

    d.add_paragraph()
    if r.night_rates:
        p = d.add_paragraph()
        para(p, "Room 1: " + r.room + f" ({r.adults}-{r.children}-0)", True)
        # Fecha | Precio | Descuento | Neto, una fila por noche; legible con n noches variables.
        tab = d.add_table(rows=1, cols=4)
        tab.style = "Table Grid"
        tab.alignment = WD_TABLE_ALIGNMENT.CENTER
        for cell, title in zip(tab.rows[0].cells, ["Fecha", "Tarifa", "Descuento", "Total noche"]):
            set_cell_shading(cell, "2261A8")
            para(cell.paragraphs[0], title, True, color=(255,255,255), size=8)
        for night in r.night_rates:
            cells = tab.add_row().cells
            for cell, content in zip(cells, [night.date, fmt_usd(night.original), "-" + fmt_usd(night.discount), fmt_usd(night.net)]):
                para(cell.paragraphs[0], content, size=8)
        p = d.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        para(p, "Total: " + fmt_usd(r.total), True)

    p = d.add_paragraph()
    para(p, f"Number of Nights: {r.nights}\nCheck-In Date: {fmt_date(r.checkin)}\nCheck-Out Date: {fmt_date(r.checkout)}")
    p = d.add_paragraph()
    para(p, "Rooms", True, size=10)
    d.add_paragraph(f"Room 1: {r.adults} Adults, {r.children} Children")
    for label, val in [
        ("Room Selected", r.room), ("Selected Board Plan", r.meal_plan),
        ("Rate", r.rate), ("Promotion", r.promotion),
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
    try:
        booking = parse_paraty_eml(mail_data)
    except Exception as exc:
        st.error(f"No se pudo interpretar este correo como confirmación Paraty: {exc}")
        st.stop()

    if booking.warnings:
        for warning in booking.warnings:
            st.warning(warning)
    st.subheader(f"Paraty: {booking.reference}")
    col1, col2, col3 = st.columns(3)
    col1.metric("Importe", fmt_usd(booking.total))
    col2.metric("Noches", booking.nights)
    col3.metric("Habitación", booking.room)
    st.write("**Huésped:**", booking.name)
    st.write("**Estancia:**", fmt_date(booking.checkin), "—", fmt_date(booking.checkout))
    if booking.night_rates:
        st.dataframe([{"Fecha": n.date, "Tarifa USD": float(n.original),
                      "Descuento USD": float(n.discount), "Neto USD": float(n.net)}
                     for n in booking.night_rates], hide_index=True, use_container_width=True)
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
        tc_input = st.number_input("Tipo de cambio del hotel (MXN por USD)",
                                   min_value=0.0, value=0.0, step=0.01, format="%.2f")
        if tc_input > 0:
            tc = Decimal(str(tc_input))
            mxn = (booking.total * tc).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            st.success(f"Equivalente: MXN ${mxn:,.2f}")
        else:
            st.warning("Captura el tipo de cambio para generar el documento American Express.")
    last_four = st.text_input("Últimos cuatro dígitos (opcional)",
                              max_chars=4, key="last_four")
    bad_four = bool(last_four and (not last_four.isdigit() or len(last_four) != 4))
    if bad_four:
        st.error("Escribe exactamente cuatro dígitos o deja el campo vacío.")

    st.subheader("Imagen de referencia (esquina superior derecha)")
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
    uploaded_image = st.file_uploader("O subir imagen ya enmascarada", type=["png", "jpg", "jpeg"], key=f"image_upload_{source_key}_{st.session_state.get('image_upload_version', 0)}")
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
        st.image(payment_image, caption="Vista previa: imagen que quedará en la parte superior derecha", width=280)
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

    disabled = (card_type == "American Express" and tc is None) or bad_four or (bool(payment_image) and not acknowledge)
    if not disabled:
        try:
            output = generate_docx(booking, card_type, tc, last_four,
                                   inssist_number=inssist_number, payment_image=payment_image)
            st.download_button("Descargar confirmación Word", output,
                               file_name=f"Confirmacion_{booking.reference}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               type="primary")
        except Exception as exc:
            st.error(f"No se pudo generar el Word: {exc}")
    else:
        st.info("Completa los datos requeridos antes de descargar el documento.")


if __name__ == "__main__":
    main()
