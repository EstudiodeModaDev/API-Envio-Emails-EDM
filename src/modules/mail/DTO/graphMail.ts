
import {ArrayMinSize, IsArray, IsBoolean, IsEmail, IsIn, IsNotEmpty, IsOptional, IsString, ValidateNested,} from 'class-validator';
import { Type } from 'class-transformer';

export class EmailAddressDto {
  @IsEmail()
  address!: string;
}

export class GraphRecipientDto {
  @ValidateNested()
  @Type(() => EmailAddressDto)
  emailAddress!: EmailAddressDto;
}

export class GraphAttachmentDto {
  @IsString()
  @IsNotEmpty()
  '@odata.type': '#microsoft.graph.fileAttachment';

  @IsString()
  @IsNotEmpty()
  name!: string;

  @IsOptional()
  @IsString()
  contentType?: string;

  @IsOptional()
  @IsString()
  contentBytes?: string;
}

export class GraphBodyDto {
  @IsIn(['Text', 'HTML'])
  contentType!: 'Text' | 'HTML';

  @IsString()
  @IsNotEmpty()
  content!: string;
}

export class GraphMessageDto {
  @IsString()
  @IsNotEmpty()
  subject!: string;

  @ValidateNested()
  @Type(() => GraphBodyDto)
  body!: GraphBodyDto;

  @IsArray()
  @ArrayMinSize(1)
  @ValidateNested({ each: true })
  @Type(() => GraphRecipientDto)
  toRecipients!: GraphRecipientDto[];

  @IsOptional()
  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => GraphRecipientDto)
  ccRecipients?: GraphRecipientDto[];

  @IsOptional()
  @IsArray()
  @ValidateNested({ each: true })
  @Type(() => GraphAttachmentDto)
  attachments?: GraphAttachmentDto[];
}

export class SendMailRequestDto {
  @IsEmail()
  senderMail!: string;

  @ValidateNested()
  @Type(() => GraphMessageDto)
  message!: GraphMessageDto;

  @IsOptional()
  @IsBoolean()
  saveToSentItems?: boolean;
}
