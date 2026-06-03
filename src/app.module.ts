import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { MailModule } from './modules/mail/mail.module';
import { PublicModule } from './modules/public/public.module';


@Module({
  imports: [
    ConfigModule.forRoot({ isGlobal: true }),
    MailModule,
    PublicModule,
  ],
})
export class AppModule {}
