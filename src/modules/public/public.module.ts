import { Module } from '@nestjs/common';
import { MailModule } from '../mail/mail.module';
import { PublicController } from './public.controller';

@Module({
  imports: [MailModule],
  controllers: [PublicController],
})
export class PublicModule {}
