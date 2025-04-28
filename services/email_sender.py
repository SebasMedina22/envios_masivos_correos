# services/email_sender.py

from email.message import EmailMessage
from email.utils import make_msgid

def preparar_mensaje(correo_emisor, correo_destino, asunto, cuerpo):
    """
    Prepara un mensaje de correo electrónico con cuerpo HTML y una imagen embebida.

    Args:
        correo_emisor (str): Dirección de correo de quien envía.
        correo_destino (str): Dirección de correo del destinatario.
        asunto (str): Asunto del correo.
        cuerpo (str): Plantilla HTML del correo, con el marcador cid:logo_astrana.

    Returns:
        EmailMessage: Mensaje de correo listo para enviar.
    """
    mensaje = EmailMessage()
    mensaje['Subject'] = asunto
    mensaje['From'] = correo_emisor
    mensaje['To'] = correo_destino

    # Cargar imagen y preparar el contenido HTML con el logo
    with open('images/logo_astrana.png', 'rb') as img:
        logo_cid = make_msgid(domain='astrana.com')
        cuerpo_final = cuerpo.replace('cid:logo_astrana', f'cid:{logo_cid[1:-1]}')

        # Agregar el cuerpo HTML
        mensaje.add_alternative(cuerpo_final, subtype='html')

        # Adjuntar imagen embebida
        mensaje.get_payload()[0].add_related(img.read(), maintype='image', subtype='png', cid=logo_cid)

    return mensaje






