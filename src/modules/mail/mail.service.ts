import {
  BadRequestException,
  ForbiddenException,
  Injectable,
  ServiceUnavailableException,
} from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { SendMailRequestDto } from './DTO/graphMail';

@Injectable()
export class MailService {
  constructor(private config: ConfigService) {}

  async sendMail({ senderMail, ...mailData }: SendMailRequestDto) {
    if (!mailData) {
      throw new BadRequestException(
        'No se proporciono informacion de correo valida',
      );
    }

    if (!senderMail) {
      throw new BadRequestException(
        'No se proporciono el correo del remitente',
      );
    }

    const normalizedSender = senderMail.trim().toLowerCase();
    const graphClient = await this.createGraphClient();
    const isAllowed = normalizedSender
      ? this.isAllowedMail(normalizedSender)
      : true;

    if (!isAllowed) {
      throw new ForbiddenException('El correo del remitente no esta permitido');
    }

    try {
      return await graphClient.api(`/users/${normalizedSender}/sendMail`).post(mailData);
    } catch {
      throw new BadRequestException('Error al enviar correo');
    }
  }

  isAllowedMail(mail: string): boolean {
    const allowedMails = this.config
      .getOrThrow('ALLOWED_MAILS')
      .split(',')
      .map((m) => m.trim().toLowerCase());

    return allowedMails.includes(mail.trim().toLowerCase());
  }

  async testConnection() {
    const mailbox = this.config
      .getOrThrow<string>('MAIL_FROM')
      .trim()
      .toLowerCase();

    try {
      await this.createGraphClient();

      return {
        ok: true,
        message: 'Autenticacion con Microsoft Graph verificada correctamente',
        mailbox,
      };
    } catch {
      throw new ServiceUnavailableException(
        'No fue posible autenticar contra Microsoft Graph',
      );
    }
  }

  private async createGraphClient() {
    const credential = new ClientSecretCredential(
      this.config.getOrThrow('TENANT_ID'),
      this.config.getOrThrow('CLIENT_ID'),
      this.config.getOrThrow('CLIENT_SECRET'),
    );

    const token = await credential.getToken(
      'https://graph.microsoft.com/.default',
    );

    if (!token?.token) {
      throw new ServiceUnavailableException(
        'No fue posible obtener un token de acceso para Microsoft Graph',
      );
    }

    return Client.init({
      authProvider: (done) => {
        done(null, token.token);
      },
    });
  }
}
