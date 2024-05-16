package de.fricofresh.outlookspitter.utils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hsmf.exceptions.ChunkNotFoundException;

import ch.astorm.jotlmsg.OutlookMessageAttachment;
import ch.astorm.jotlmsg.OutlookMessageRecipient;
import de.fricofresh.outlookspitter.CreateSplittedFilesParameter;
import jakarta.activation.DataHandler;
import jakarta.mail.Flags.Flag;
import jakarta.mail.Message;
import jakarta.mail.MessagingException;
import jakarta.mail.internet.MimeBodyPart;
import jakarta.mail.internet.MimeMessage;
import jakarta.mail.internet.MimeMessage.RecipientType;
import jakarta.mail.internet.MimeMultipart;
import jakarta.mail.util.ByteArrayDataSource;

public class JavaMailMessageUtil {
	
	private static Logger log = LogManager.getLogger(JavaMailMessageUtil.class);
	
	public static Message createMessage(CreateSplittedFilesParameter parameters) {
		
		try {
			OutlookMessageExtended outlookMessageExtended = new OutlookMessageExtended(parameters.getEmailPath().toFile());
			MimeMessage msg = outlookMessageExtended.toMimeMessage();
			
			if (parameters.getEmailHTMLMessage().isPresent()) {
				byte[] encoded = Files.readAllBytes(Paths.get(parameters.getEmailHTMLMessage().get()));
				String htmlContent = new String(encoded, StandardCharsets.ISO_8859_1);
				// FIXME
				htmlContent = Pattern.compile("(\\r\\n.+)+\\s{7,}.+(\\r\\n.+)+", Pattern.MULTILINE).matcher(htmlContent).replaceAll("");
				MimeBodyPart body = new MimeBodyPart();
				MimeMultipart multipart = new MimeMultipart();
				body.setText(htmlContent, "UTF-8", "html");
				multipart.addBodyPart(body);
				try {
					for (OutlookMessageAttachment attachment : outlookMessageExtended.getAttachments()) {
						body = new MimeBodyPart();
						ByteArrayDataSource byteArrayDataSource = new ByteArrayDataSource(attachment.getNewInputStream(), attachment.getMimeType());
						body.setDataHandler(new DataHandler(byteArrayDataSource));
						body.setFileName(attachment.getName());
						multipart.addBodyPart(body);
					}
				}
				catch (IOException | MessagingException e) {
					log.error(e, e);
				}
				msg.setContent(multipart);
			}
			
			msg.setSentDate(null);
			msg.setReplyTo(null);
			msg.removeHeader("Date");
			msg.setHeader("X-Unsent", "1");
			msg.setFlag(Flag.DRAFT, false);
			
			return msg;
		}
		catch (IOException | MessagingException e) {
			log.error(e, e);
		}
		return null;
	}
	
	public static List<Path> createSplittedFiles(CreateSplittedFilesParameter parameterObject) throws IOException, ChunkNotFoundException, MessagingException {
		
		List<Path> result = new ArrayList<>();
		
		Message tempMessage = createMessage(parameterObject);
		
		for (int i = 0; i < parameterObject.getRecipientsToSplit().size(); i++) {
			OutlookMessageRecipient recipient = parameterObject.getRecipientsToSplit().get(i);
			try {
				if (i != 0 && i % parameterObject.getSplit() == 0) {
					tempMessage = writeTempMessage(parameterObject, result, tempMessage);
				}
				tempMessage.addRecipient(convertOutlookRecipientsToJakarta(recipient), recipient.getAddress());
			}
			catch (MessagingException e) {
				log.error(e, e);
			}
		}
		writeTempMessage(parameterObject, result, tempMessage);
		
		return result;
	}
	
	public static Message.RecipientType convertOutlookRecipientsToJakarta(OutlookMessageRecipient outlookMessageRecipient) {
		
		Message.RecipientType result;
		
		switch (outlookMessageRecipient.getType()) {
			case TO:
				result = RecipientType.TO;
				break;
			case CC:
				result = RecipientType.CC;
				break;
			case BCC:
				result = RecipientType.BCC;
				break;
			default:
				throw new RuntimeException();
		}
		
		return result;
	}
	
	public static Message writeTempMessage(CreateSplittedFilesParameter parameterObject, List<Path> result, Message tempMessage) throws IOException, ChunkNotFoundException, MessagingException {
		
		Path outputFile = MailSplitterUtil.getOutputFile(parameterObject, result.size());
		try (FileOutputStream fos = new FileOutputStream(outputFile.toFile())) {
			
			tempMessage.writeTo(fos);
		}
		result.add(outputFile);
		
		tempMessage = createMessage(parameterObject);
		return tempMessage;
	}
	
}
