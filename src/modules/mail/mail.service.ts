import {
  BadRequestException,
  ForbiddenException,
  Injectable,
  NotFoundException,
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
      const draftMessage = await graphClient
        .api(`/users/${normalizedSender}/messages`)
        .post(mailData.message);

      await graphClient
        .api(`/users/${normalizedSender}/messages/${draftMessage.id}/send`)
        .post({});

      return {
        ok: true,
        message: 'Correo enviado correctamente',
        graph: {
          id: draftMessage.id,
          internetMessageId: draftMessage.internetMessageId ?? null,
          conversationId: draftMessage.conversationId ?? null,
          webLink: draftMessage.webLink ?? null,
          createdDateTime: draftMessage.createdDateTime ?? null,
          lastModifiedDateTime: draftMessage.lastModifiedDateTime ?? null,
          subject: draftMessage.subject ?? mailData.message.subject,
        },
      };
    } catch (error: any) {
      this.handleGraphError(
        error,
        `No fue posible enviar el correo usando el buzon ${normalizedSender}`,
      );
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
      const graphClient = await this.createGraphClient();
      const mailboxInfo = await graphClient
        .api(`/users/${mailbox}?$select=id,mail,userPrincipalName`)
        .get();

      return {
        ok: true,
        message:
          'Autenticacion y acceso al buzon verificados correctamente en Microsoft Graph',
        mailbox,
        mailboxInfo: {
          id: mailboxInfo.id ?? null,
          mail: mailboxInfo.mail ?? null,
          userPrincipalName: mailboxInfo.userPrincipalName ?? null,
        },
      };
    } catch (error: any) {
      this.handleGraphError(
        error,
        `No fue posible validar el acceso al buzon ${mailbox}`,
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

  private handleGraphError(error: any, fallbackMessage: string): never {
    const statusCode = error?.statusCode ?? error?.code ?? 500;
    const graphCode = error?.body?.error?.code ?? error?.code ?? null;
    const graphMessage = error?.body?.error?.message ?? error?.message ?? null;

    if (
      statusCode === 403 ||
      graphCode === 'ErrorAccessDenied' ||
      graphMessage?.toLowerCase().includes('access denied')
    ) {
      throw new ForbiddenException({
        message: fallbackMessage,
        detail:
          'Microsoft Graph rechazo la operacion por permisos insuficientes sobre el buzon. Verifica que la app tenga permisos de aplicacion Mail.ReadWrite y Mail.Send con admin consent, y que el buzon exista y este habilitado.',
        graph: {
          code: graphCode,
          message: graphMessage,
        },
      });
    }

    if (statusCode === 404 || graphCode === 'ResourceNotFound') {
      throw new NotFoundException({
        message: fallbackMessage,
        detail:
          'El buzon indicado no existe en Microsoft 365 o no es accesible con el identificador enviado.',
        graph: {
          code: graphCode,
          message: graphMessage,
        },
      });
    }

    if (statusCode === 401 || graphCode === 'InvalidAuthenticationToken') {
      throw new ServiceUnavailableException({
        message: fallbackMessage,
        detail:
          'Microsoft Graph rechazo el token de acceso. Revisa TENANT_ID, CLIENT_ID, CLIENT_SECRET y el admin consent de la aplicacion.',
        graph: {
          code: graphCode,
          message: graphMessage,
        },
      });
    }

    throw new BadRequestException({
      message: fallbackMessage,
      detail: graphMessage ?? 'Error no controlado al invocar Microsoft Graph',
      graph: {
        code: graphCode,
        message: graphMessage,
      },
    });
  }
}
