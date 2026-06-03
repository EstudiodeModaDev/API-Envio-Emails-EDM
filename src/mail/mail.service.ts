import { Injectable } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';

@Injectable()
export class MailService {
  constructor(private config: ConfigService) {}

  async sendMail(form: any) {
    const credential = new ClientSecretCredential(
      this.config.getOrThrow('TENANT_ID'),
      this.config.getOrThrow('CLIENT_ID'),
      this.config.getOrThrow('CLIENT_SECRET'),
    );

    const token = await credential.getToken(
      'https://graph.microsoft.com/.default',
    );

    console.log(token)

    const graphClient = Client.init({
      authProvider: (done) => {
        done(null, token?.token ?? '');
      },
    });

    const mailData = {
      message: {
        subject: `Nueva denuncia recibida (${form.correo})`,
        body: {
          contentType: 'HTML',
          content: `
            <p>Se ha recibido una nueva denuncia:</p>
            <ul>
              <li><strong>Cédula:</strong> ${form.cedula || 'No proporcionada'}</li>
              <li><strong>Teléfono:</strong> ${form.telefono || 'No proporcionado'}</li>
              <li><strong>Correo:</strong> ${form.correo}</li>
              <li><strong>Empleado EDM:</strong> ${form.esEmpleado ? 'Sí' : 'No'}</li>
              <li><strong>Denuncia:</strong> ${form.denuncia}</li>
            </ul>
          `,
        },
        toRecipients: [
          {
            emailAddress: {
              address: this.config.getOrThrow('MAIL_TO'),
            },
          },
        ],
      },
      saveToSentItems: true,
    };

    try {
      const res = await graphClient
      .api(`/users/${this.config.getOrThrow('MAIL_FROM')}/sendMail`)
      .post(mailData);
      console.log('Correo enviado:', res);
      return res;
    } catch (error) {
      console.error('Error al enviar correo:', error);
      throw error;
    } 
  }
}