package de.fricofresh.outlookspitter;

import java.nio.file.Path;
import java.util.List;
import java.util.Optional;

import org.apache.poi.hsmf.MAPIMessage;

import ch.astorm.jotlmsg.OutlookMessageRecipient;

public class CreateSplittedFilesParameter {
	
	private List<OutlookMessageRecipient> recipients;
	
	private List<OutlookMessageRecipient> recipientsToSplit;
	
	private MAPIMessage emailMessage;
	
	private Path emailPath;
	
	private int split;
	
	private Optional<String> prefix = Optional.empty();
	
	private Optional<String> suffix = Optional.empty();
	
	private Optional<String> outputDir = Optional.empty();
	
	private Optional<String> emailHTMLTextPath = Optional.empty();
	
	private MailGenMethod mailGenMehtod = MailGenMethod.POI;
	
	public CreateSplittedFilesParameter() {
		
	}
	
	public CreateSplittedFilesParameter(List<OutlookMessageRecipient> recipients, MAPIMessage emailMessage, int split, Optional<String> prefix, Optional<String> suffix, Optional<String> outputDir) {
		
		this.recipients = recipients;
		this.emailMessage = emailMessage;
		this.split = split;
		this.prefix = prefix;
		this.suffix = suffix;
		this.outputDir = outputDir;
	}
	
	public List<OutlookMessageRecipient> getRecipients() {
		
		return recipients;
	}
	
	public void setRecipients(List<OutlookMessageRecipient> recipients) {
		
		this.recipients = recipients;
	}
	
	public List<OutlookMessageRecipient> getRecipientsToSplit() {
		
		return recipientsToSplit;
	}
	
	public void setRecipientsToSplit(List<OutlookMessageRecipient> recipientsToSplit) {
		
		this.recipientsToSplit = recipientsToSplit;
	}
	
	public MAPIMessage getEmailMessage() {
		
		return emailMessage;
	}
	
	public void setEmailMessage(MAPIMessage emailMessage) {
		
		this.emailMessage = emailMessage;
	}
	
	public int getSplit() {
		
		return split;
	}
	
	public void setSplit(int split) {
		
		this.split = split;
	}
	
	public Optional<String> getPrefix() {
		
		return prefix;
	}
	
	public void setPrefix(Optional<String> prefix) {
		
		this.prefix = prefix;
	}
	
	public Optional<String> getSuffix() {
		
		return suffix;
	}
	
	public void setSuffix(Optional<String> suffix) {
		
		this.suffix = suffix;
	}
	
	public Optional<String> getOutputDir() {
		
		return outputDir;
	}
	
	public void setOutputDir(Optional<String> outputDir) {
		
		this.outputDir = outputDir;
	}
	
	public MailGenMethod getMailGenMehtod() {
		
		return mailGenMehtod;
	}
	
	public void setMailGenMehtod(MailGenMethod mailGenMehtod) {
		
		this.mailGenMehtod = mailGenMehtod;
	}
	
	public Path getEmailPath() {
		
		return emailPath;
	}
	
	public void setEmailPath(Path emailPath) {
		
		this.emailPath = emailPath;
	}
	
	public Optional<String> getEmailHTMLMessage() {
		
		return emailHTMLTextPath;
	}
	
	public void setEmailHTMLMessage(Optional<String> emailHTMLMessage) {
		
		this.emailHTMLTextPath = emailHTMLMessage;
	}
}
