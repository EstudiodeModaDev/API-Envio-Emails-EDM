# Mail API

API en NestJS para envio de correos a traves de Microsoft Graph.

## Requisitos

- Node.js 18 o superior
- Una aplicacion registrada en Azure AD con permisos para enviar correo con Microsoft Graph
- Uno o varios buzones autorizados en la variable `ALLOWED_MAILS`

## Variables de entorno

Crea un archivo `.env` con estas variables:

```env
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
PORT=4141
ALLOWED_MAILS=alert@dominio.com,soporte@dominio.com
```

Descripcion:

- `TENANT_ID`: tenant de Azure AD.
- `CLIENT_ID`: application/client ID de la app registrada.
- `CLIENT_SECRET`: secreto de la aplicacion.
- `PORT`: puerto del servidor.
- `ALLOWED_MAILS`: lista separada por comas con los buzones desde los que esta permitido enviar.

## Instalacion

```bash
npm install
```

## Ejecucion

```bash
# desarrollo
npm run start

# watch mode
npm run start:dev

# produccion
npm run start:prod
```

Por defecto la API escucha en `http://localhost:4141` si `PORT=4141`.

## Endpoint

### `POST /mail/send`

Envia un correo usando el buzón indicado en `senderMail`, siempre que ese correo exista en `ALLOWED_MAILS`.

#### Headers

```http
Content-Type: application/json
```

#### Body

```json
{
  "senderMail": "alert@dominio.com",
  "message": {
    "subject": "Prueba de correo",
    "body": {
      "contentType": "HTML",
      "content": "<p>Hola, este es un correo de prueba.</p>"
    },
    "toRecipients": [
      {
        "emailAddress": {
          "address": "destinatario@dominio.com"
        }
      }
    ],
    "ccRecipients": [
      {
        "emailAddress": {
          "address": "copia@dominio.com"
        }
      }
    ],
    "attachments": [
      {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": "archivo.txt",
        "contentType": "text/plain",
        "contentBytes": "SG9sYSBtdW5kbw=="
      }
    ]
  },
  "saveToSentItems": true
}
```

#### Validaciones

La API valida el request con `ValidationPipe`, `class-validator` y `class-transformer`.

Reglas principales:

- `senderMail` debe ser un correo valido.
- `senderMail` debe existir en `ALLOWED_MAILS`.
- `message.subject` es obligatorio.
- `message.body.contentType` solo permite `Text` o `HTML`.
- `message.body.content` es obligatorio.
- `message.toRecipients` debe tener al menos un destinatario.
- Cada `emailAddress.address` debe ser un correo valido.
- `saveToSentItems` es opcional y debe ser booleano.

Si se envian propiedades no definidas en el DTO, la API las rechaza.

#### Respuesta exitosa

```json
{
  "ok": true,
  "message": "Correo enviado correctamente"
}
```

#### Errores esperados

- `400 Bad Request`: body invalido o campos faltantes.
- `403 Forbidden`: `senderMail` no autorizado.
- `500 Internal Server Error`: error al enviar con Microsoft Graph o problema interno del servidor.

## Ejemplo con cURL

```bash
curl -X POST http://localhost:4141/mail/send \
  -H "Content-Type: application/json" \
  -d '{
    "senderMail": "alert@dominio.com",
    "message": {
      "subject": "Correo desde la API",
      "body": {
        "contentType": "HTML",
        "content": "<p>Mensaje de prueba</p>"
      },
      "toRecipients": [
        {
          "emailAddress": {
            "address": "destinatario@dominio.com"
          }
        }
      ]
    },
    "saveToSentItems": true
  }'
```

## Notas

- `contentBytes` debe enviarse en Base64 cuando adjuntes archivos.
- El envio depende de que la aplicacion de Azure tenga permisos correctos sobre Microsoft Graph.
- Aunque la API permite varios buzones, solo enviara desde correos incluidos en `ALLOWED_MAILS`.
