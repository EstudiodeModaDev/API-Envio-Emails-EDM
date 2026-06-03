import { BadRequestException, ForbiddenException, Injectable } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { SendMailRequestDto } from './DTO/graphMail';

@Injectable()
export class MailService {
  constructor(private config: ConfigService) {}

  async sendMail({senderMail, ...mailData}: SendMailRequestDto,) {
    
    
    if(!mailData){
      throw new BadRequestException('No se proporcionó información de correo válida');
    }

    if(!senderMail){
      throw new BadRequestException('No se proporcionó el correo del remitente');
    }

    const normalizedSender = senderMail.trim().toLowerCase();

    const credential = new ClientSecretCredential(
      this.config.getOrThrow('TENANT_ID'),
      this.config.getOrThrow('CLIENT_ID'),
      this.config.getOrThrow('CLIENT_SECRET'),
    );

    const token = await credential.getToken(
      'https://graph.microsoft.com/.default',
    );

    const graphClient = Client.init({
      authProvider: (done) => {
        done(null, token?.token ?? '');
      },
    });

    const isAllowed = normalizedSender ? this.isAllowedMail(normalizedSender) : true;

    if (!isAllowed) {
      throw new ForbiddenException('El correo del remitente no está permitido');
    }

    try {
      const res = await graphClient
      .api(`/users/${normalizedSender}/sendMail`)
      .post(mailData);
      return res;
    } catch (error) {
      throw new BadRequestException('Error al enviar correo');
    } 
  }

  isAllowedMail(mail: string): boolean {
    const allowedMails = this.config.getOrThrow('ALLOWED_MAILS').split(',').map((m) => m.trim().toLowerCase());
    return allowedMails.includes(mail.trim().toLowerCase());
  }

}