package de.fricofresh.outlookspitter.utils;

import static ch.astorm.jotlmsg.io.PropertiesChunk.FLAG_READABLE;
import static ch.astorm.jotlmsg.io.PropertiesChunk.FLAG_WRITEABLE;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.ByteBuffer;
import java.nio.file.Path;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

import org.apache.poi.hsmf.MAPIMessage;
import org.apache.poi.hsmf.datatypes.AttachmentChunks;
import org.apache.poi.hsmf.datatypes.MAPIProperty;
import org.apache.poi.hsmf.datatypes.NameIdChunks;
import org.apache.poi.hsmf.datatypes.PropertyValue;
import org.apache.poi.hsmf.datatypes.RecipientChunks;
import org.apache.poi.hsmf.datatypes.Types;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.StringUtil;

import ch.astorm.jotlmsg.OutlookMessage;
import ch.astorm.jotlmsg.OutlookMessageAttachment;
import ch.astorm.jotlmsg.OutlookMessageRecipient;
import ch.astorm.jotlmsg.OutlookMessageRecipient.Type;
import ch.astorm.jotlmsg.io.FlatEntryListStructure;
import ch.astorm.jotlmsg.io.MessagePropertiesChunk;
import ch.astorm.jotlmsg.io.OneOffEntryIDStructure;
import ch.astorm.jotlmsg.io.PropertiesChunk;
import ch.astorm.jotlmsg.io.StoragePropertiesChunk;

public class OutlookMessageExtended extends OutlookMessage {
	
	public OutlookMessageExtended() {
		
		super();
	}
	
	public OutlookMessageExtended(File mapiMessageFile) throws IOException {
		
		super(mapiMessageFile);
	}
	
	public OutlookMessageExtended(InputStream mapiMessageInputStream) throws IOException {
		
		super(mapiMessageInputStream);
	}
	
	public OutlookMessageExtended(MAPIMessage mapiMessage) {
		
		super(mapiMessage);
	}
	
	public void setEMailRecipientTo(List<String> emailAddresses) {
		
		clearEMailRecipient(Type.TO);
		emailAddresses.stream().map(emailAddress -> new OutlookMessageRecipient(Type.TO, emailAddress)).forEach(this::addRecipient);
	}
	
	public void setEMailRecipientCC(List<String> emailAddresses) {
		
		clearEMailRecipient(Type.CC);
		emailAddresses.stream().map(emailAddress -> new OutlookMessageRecipient(Type.TO, emailAddress)).forEach(this::addRecipient);
	}
	
	public void setEMailRecipientBCC(List<String> emailAddresses) {
		
		clearEMailRecipient(Type.BCC);
		emailAddresses.stream().map(emailAddress -> new OutlookMessageRecipient(Type.TO, emailAddress)).forEach(this::addRecipient);
	}
	
	public void setEMailAllRecipient(List<OutlookMessageRecipient> emailAddresses) {
		
		removeAllRecipients();
		emailAddresses.stream().forEach(this::addRecipient);
	}
	
	public void clearEMailRecipient() {
		
		removeAllRecipients();
	}
	
	public void clearEMailRecipient(Type type) {
		
		removeAllRecipients(type);
	}
	
	public static List<String> splitAdresses(String adresses) {
		
		return Arrays.asList(adresses.split(";")).stream().map(OutlookMessageExtended::extractEmail).collect(Collectors.toList());
	}
	
	public void writeTo(Path pathToFile) throws IOException {
		
		writeTo(pathToFile.toFile());
	}
	
	public void writeAndOpen(Path pathToFile) throws IOException {
		
		writeTo(pathToFile.toFile());
		MailSplitterUtil.openFiles(Arrays.asList(pathToFile), Optional.empty());
	}
	
	/**
	 * Writes the content of this message to the specified {@code outputStream}. The bytes written represent a {@code .msg} file that can be open by Microsoft Outlook. The {@code outputStream} will remain open.
	 * 
	 * @param file
	 *            The file to write to.
	 * @throws IOException
	 *             If an I/O error occurs.
	 */
	public void writeToEditiable(File file) throws IOException {
		
		FileOutputStream outputStream = new FileOutputStream(file);
		POIFSFileSystem fs = new POIFSFileSystem();
		
		List<OutlookMessageRecipient> recipients = getAllRecipients();
		List<OutlookMessageAttachment> attachments = getAttachments();
		List<String> replyToRecipents = getReplyTo();
		String body = getPlainTextBody();
		String subject = getSubject();
		String from = getFrom();
		
		// creates the basic structure (page 17, point 2.2.3)
		DirectoryEntry nameid = fs.createDirectory(NameIdChunks.NAME);
		nameid.createDocument(PropertiesChunk.PREFIX + "00020102", new ByteArrayInputStream(new byte[0])); // GUID Stream
		nameid.createDocument(PropertiesChunk.PREFIX + "00030102", new ByteArrayInputStream(new byte[0])); // Entry Stream (mandatory, otherwise Outlook crashes)
		nameid.createDocument(PropertiesChunk.PREFIX + "00040102", new ByteArrayInputStream(new byte[0])); // String Stream
		
		// creates the top-level structure of data
		MessagePropertiesChunk topLevelChunk = new MessagePropertiesChunk();
		topLevelChunk.setAttachmentCount(attachments.size());
		topLevelChunk.setRecipientCount(recipients.size());
		topLevelChunk.setNextAttachmentId(attachments.size()); // actually indicates the next free id !
		topLevelChunk.setNextRecipientId(recipients.size()); // actually indicates the next free id !
		
		// constants values can be found here: https://msdn.microsoft.com/en-us/library/ee219881(v=exchg.80).aspx
		topLevelChunk.setProperty(new PropertyValue(MAPIProperty.STORE_SUPPORT_MASK, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(0x00040000).array())); // all the strings will be in unicode
		topLevelChunk.setProperty(new PropertyValue(MAPIProperty.MESSAGE_CLASS, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE("IPM.Note"))); // outlook message
		topLevelChunk.setProperty(new PropertyValue(MAPIProperty.HASATTACH, FLAG_READABLE | FLAG_WRITEABLE, attachments.isEmpty() ? new byte[] {0} : new byte[] {1}));
		
		if (getSentDate() == null) {
			topLevelChunk.setProperty(new PropertyValue(MAPIProperty.MESSAGE_FLAGS, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(0x00000008).array()));
		} // mfUnsent - https://msdn.microsoft.com/en-us/library/ee160304(v=exchg.80).aspx
		else {
			Method methodDateToBytes;
			try {
				methodDateToBytes = getClass().getSuperclass().getDeclaredMethod("dateToBytes", new Class<?>[] {Date.class});
				methodDateToBytes.setAccessible(true);
				SimpleDateFormat mdf = new SimpleDateFormat(MIME_DATE_FORMAT);
				topLevelChunk.setProperty(new PropertyValue(MAPIProperty.MESSAGE_FLAGS, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(2).array())); // mfUnmodified
				topLevelChunk.setProperty(new PropertyValue(MAPIProperty.CLIENT_SUBMIT_TIME, FLAG_READABLE | FLAG_WRITEABLE, (byte[]) methodDateToBytes.invoke(this, getSentDate())));
				topLevelChunk.setProperty(new PropertyValue(MAPIProperty.TRANSPORT_MESSAGE_HEADERS, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE("Date: " + mdf.format(getSentDate()))));
			}
			catch (NoSuchMethodException | SecurityException | IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
				e.printStackTrace();
			}
		}
		// set it editable
		topLevelChunk.setProperty(new PropertyValue(MAPIProperty.ACCESS, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(0x00000001).array()));
		topLevelChunk.setProperty(new PropertyValue(MAPIProperty.ACCESS_LEVEL, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(0x00000001).array()));
		topLevelChunk.setProperty(new PropertyValue(MAPIProperty.createCustom(0x001F, Types.ASCII_STRING, from), FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE("newMail")));
		
		if (subject != null) {
			topLevelChunk.setProperty(new PropertyValue(MAPIProperty.SUBJECT, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(subject)));
		}
		if (body != null) {
			topLevelChunk.setProperty(new PropertyValue(MAPIProperty.BODY, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(body)));
		}
		if (from != null) {
			topLevelChunk.setProperty(new PropertyValue(MAPIProperty.SENDER_EMAIL_ADDRESS, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(from)));
			topLevelChunk.setProperty(new PropertyValue(MAPIProperty.SENDER_NAME, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(from)));
		}
		
		// creates the reply recipients
		if (replyToRecipents != null) {
			FlatEntryListStructure<OneOffEntryIDStructure> fels = new FlatEntryListStructure<>();
			String replyToRecipentNames = null;
			boolean first = true;
			for (String replyToRecipent : replyToRecipents) {
				if (first) {
					replyToRecipentNames = new String();
					first = false;
				}
				else {
					replyToRecipentNames += ";";
				}
				replyToRecipentNames += replyToRecipent;
				fels.addFlatEntryStructure(new OneOffEntryIDStructure(replyToRecipent));
			}
			
			if (replyToRecipentNames != null && fels.getCount() > 0) {
				// Note: There must be responding REPLY_RECIPIENT_ENTRIES and REPLY_RECIPIENT_NAMES MAPIProperties.
				topLevelChunk.setProperty(new PropertyValue(MAPIProperty.REPLY_RECIPIENT_ENTRIES, FLAG_READABLE | FLAG_WRITEABLE, fels.toBytes()));
				topLevelChunk.setProperty(new PropertyValue(MAPIProperty.REPLY_RECIPIENT_NAMES, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(replyToRecipentNames)));
			}
		}
		
		topLevelChunk.writeTo(fs.getRoot());
		
		// creates the recipients
		int recipientCounter = 0;
		for (OutlookMessageRecipient recipient : recipients) {
			if (recipientCounter >= 2048) {
				throw new RuntimeException("too many recipients (max=2048)");
			} // limitation, see page 15, point 2.2.1
			
			String name = recipient.getName();
			String email = recipient.getEmail();
			Type type = recipient.getType();
			
			int rt = type == Type.TO ? 1 : type == Type.CC ? 2 : 3;
			
			StoragePropertiesChunk recipStorage = new StoragePropertiesChunk();
			recipStorage.setProperty(new PropertyValue(MAPIProperty.OBJECT_TYPE, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(6).array())); // MAPI_MAILUSER
			recipStorage.setProperty(new PropertyValue(MAPIProperty.DISPLAY_TYPE, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(0).array())); // DT_MAILUSER
			recipStorage.setProperty(new PropertyValue(MAPIProperty.RECIPIENT_TYPE, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(rt).array()));
			recipStorage.setProperty(new PropertyValue(MAPIProperty.ROWID, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(recipientCounter).array()));
			// set it editable
			recipStorage.setProperty(new PropertyValue(MAPIProperty.ACCESS, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(1).array()));
			recipStorage.setProperty(new PropertyValue(MAPIProperty.ACCESS_LEVEL, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(1).array()));
			if (name != null) {
				recipStorage.setProperty(new PropertyValue(MAPIProperty.DISPLAY_NAME, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(name)));
				recipStorage.setProperty(new PropertyValue(MAPIProperty.RECIPIENT_DISPLAY_NAME, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(name)));
			}
			if (email != null) {
				recipStorage.setProperty(new PropertyValue(MAPIProperty.EMAIL_ADDRESS, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(email)));
				if (name == null) {
					recipStorage.setProperty(new PropertyValue(MAPIProperty.DISPLAY_NAME, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(email)));
					recipStorage.setProperty(new PropertyValue(MAPIProperty.RECIPIENT_DISPLAY_NAME, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(email)));
				}
			}
			
			String rid = "" + Integer.toHexString(recipientCounter);
			while (rid.length() < 8) {
				rid = "0" + rid;
			}
			DirectoryEntry recip = fs.createDirectory(RecipientChunks.PREFIX + rid); // page 15, point 2.2.1
			recipStorage.writeTo(recip);
			
			++recipientCounter;
		}
		
		// creates the attachments
		int attachmentCounter = 0;
		for (OutlookMessageAttachment attachment : attachments) {
			if (attachmentCounter >= 2048) {
				throw new RuntimeException("too many attachments (max=2048)");
			} // limitation, see page 15, point 2.2.2
			
			String name = attachment.getName();
			String mimeType = attachment.getMimeType();
			
			Method methodReadAttachement;
			byte[] data = null;
			try {
				methodReadAttachement = getClass().getSuperclass().getDeclaredMethod("readAttachement", new Class<?>[] {});
				methodReadAttachement.setAccessible(true);
				data = (byte[]) methodReadAttachement.invoke(this, attachment);
			}
			catch (NoSuchMethodException | SecurityException | IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
				e.printStackTrace();
			}
			
			StoragePropertiesChunk attachStorage = new StoragePropertiesChunk();
			attachStorage.setProperty(new PropertyValue(MAPIProperty.OBJECT_TYPE, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(7).array())); // MAPI_ATTACH
			if (name != null) {
				attachStorage.setProperty(new PropertyValue(MAPIProperty.ATTACH_FILENAME, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(name)));
				attachStorage.setProperty(new PropertyValue(MAPIProperty.ATTACH_LONG_FILENAME, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(name)));
			}
			if (mimeType != null) {
				attachStorage.setProperty(new PropertyValue(MAPIProperty.ATTACH_MIME_TAG, FLAG_READABLE | FLAG_WRITEABLE, StringUtil.getToUnicodeLE(mimeType)));
			}
			attachStorage.setProperty(new PropertyValue(MAPIProperty.ATTACH_NUM, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(attachmentCounter).array()));
			attachStorage.setProperty(new PropertyValue(MAPIProperty.ATTACH_METHOD, FLAG_READABLE | FLAG_WRITEABLE, ByteBuffer.allocate(4).putInt(1).array())); // ATTACH_BY_VALUE
			attachStorage.setProperty(new PropertyValue(MAPIProperty.ATTACH_DATA, FLAG_READABLE | FLAG_WRITEABLE, data));
			
			String rid = "" + Integer.toHexString(attachmentCounter);
			while (rid.length() < 8) {
				rid = "0" + rid;
			}
			DirectoryEntry recip = fs.createDirectory(AttachmentChunks.PREFIX + rid); // page 15, point 2.2.1
			attachStorage.writeTo(recip);
			
			++attachmentCounter;
		}
		
		fs.writeFilesystem(outputStream);
		outputStream.close();
		fs.close();
	}
	
	@Override
	public Object clone() {
		
		try {
			return (OutlookMessageExtended) super.clone();
		}
		catch (CloneNotSupportedException e) {
			return cloneCopy();
		}
	}
	
	public OutlookMessageExtended cloneCopy() {
		
		OutlookMessageExtended result = new OutlookMessageExtended();
		result.setEMailAllRecipient(this.getAllRecipients());
		result.setFrom(this.getFrom());
		result.setPlainTextBody(this.getPlainTextBody());
		result.setSubject(this.getSubject());
		this.getAttachments().stream().forEach(result::addAttachment);
		return result;
	}
}
