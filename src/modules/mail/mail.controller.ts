import { Body, Controller, Post } from '@nestjs/common';
import { MailService } from './mail.service';
import { SendMailRequestDto } from './DTO/graphMail';

@Controller('mail')
export class MailController {
  constructor(private readonly mailService: MailService) {}

  @Post('send')
  async send(@Body() body: SendMailRequestDto) {
    await this.mailService.sendMail(body);

    return {
      ok: true,
      message: 'Correo enviado correctamente',
    };
  }
}
