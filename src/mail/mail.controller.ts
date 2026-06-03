import { Body, Controller, Post } from '@nestjs/common';
import { MailService } from './mail.service';

@Controller('mail')
export class MailController {
  constructor(private readonly mailService: MailService) {}

  @Post('send')
  async send(@Body() form: any) {
    await this.mailService.sendMail(form);

    return {
      ok: true,
      message: 'Correo enviado correctamente',
    };
  }
}