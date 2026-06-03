import { Controller, Get } from '@nestjs/common';
import { MailService } from '../mail/mail.service';

@Controller('public')
export class PublicController {
  constructor(private readonly mailService: MailService) {}

  @Get('connection-test')
  async testConnection() {
    return this.mailService.testConnection();
  }
}
