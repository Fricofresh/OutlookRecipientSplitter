package de.fricofresh.outlookspitter.utils;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.hsmf.MAPIMessage;

import ch.astorm.jotlmsg.OutlookMessage;
import ch.astorm.jotlmsg.OutlookMessageRecipient;
import ch.astorm.jotlmsg.OutlookMessageRecipient.Type;

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
		emailAddresses.stream().map(emailAddress -> new OutlookMessageRecipient(Type.TO, emailAddress))
				.forEach(this::addRecipient);
	}
	
	public void setEMailRecipientCC(List<String> emailAddresses) {
		
		clearEMailRecipient(Type.CC);
		emailAddresses.stream().map(emailAddress -> new OutlookMessageRecipient(Type.TO, emailAddress))
				.forEach(this::addRecipient);
	}
	
	public void setEMailRecipientBCC(List<String> emailAddresses) {
		
		clearEMailRecipient(Type.BCC);
		emailAddresses.stream().map(emailAddress -> new OutlookMessageRecipient(Type.TO, emailAddress))
				.forEach(this::addRecipient);
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
		
		return Arrays.asList(adresses.split(";")).stream().map(OutlookMessageExtended::extractEmail)
				.collect(Collectors.toList());
	}
	
	public void writeTo(Path pathToFile) throws IOException {
		
		writeTo(pathToFile.toFile());
	}
	
}
