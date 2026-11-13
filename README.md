# Confirmaciones Paraty → Word

Aplicación de Streamlit para convertir correos de confirmación de Paraty a documentos Word. Incluye la referencia de Paraty, campo opcional INSSIST, desglose por noche, tipo de cambio manual para American Express y una imagen *previamente enmascarada* en la esquina superior derecha.

## Ejecutar en Windows

Instalar Python 3.10+ y ejecutar desde esta carpeta:

```powershell
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

Para probar sin Gmail, selecciona **Archivo .eml** y carga el correo de Paraty.

## Gmail (solo para ejecución local, por ahora)

La conexión directa existente utiliza OAuth de **aplicación de escritorio**:
1. Crear un proyecto Google Cloud habilitando Gmail API y configurando la pantalla de consentimiento.
2. Crear credenciales OAuth del tipo **Desktop app**.
3. Guardar el JSON como `credentials.json` en la carpeta del programa **en la computadora local**, jamás en GitHub.
4. Ejecutar el programa localmente, elegir **Gmail (directo)** y conectar la cuenta autorizada.

Solicita acceso de solo lectura a Gmail. La aplicación guarda `token.json` **localmente**, excluido por `.gitignore`.

## Streamlit Community Cloud: precaución

Esta versión **NO está lista para Gmail directo desde Streamlit Community Cloud**: `InstalledAppFlow.run_local_server()` solo funciona cuando el navegador y la aplicación corren en el equipo local. Para una versión alojada se requiere implementar OAuth web con redirect URI HTTPS, tokens separados por usuario y acceso restringido al personal autorizado.

La app procesa información personal de huéspedes. **No publiques una instancia accesible a cualquier persona** ni configures credenciales de Gmail hasta añadir autenticación y controles de acceso. Un repositorio público es distinto de una aplicación pública; este repositorio debe contener solo código, sin datos reales, capturas ni claves.

## Imágenes de pago

La imagen se pega mediante el botón del portapapeles o se sube en PNG/JPG, y se inserta arriba a la derecha. Debe estar **enmascarada previamente**: no introducir PAN completo, CVV, fecha de vencimiento ni datos de pago sensibles. El programa no enmascara ni verifica automáticamente la imagen. Los detalles bancarios se consultan solamente en un sistema autorizado, como Paraty Manager.

## Archivos

- `app.py`: aplicación.
- `requirements.txt`: dependencias.
- `.gitignore`: exclusión de credenciales y materiales de huéspedes.
- `README.md`: esta guía.

La carga `.eml`, el campo INSSIST, la generación Word y la conversión American Express se probaron con un archivo de muestra. Gmail OAuth real requiere autorización del hotel y no se ha probado.
