package de.fricofresh.outlookspitter.utils;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hsmf.exceptions.ChunkNotFoundException;

import ch.astorm.jotlmsg.OutlookMessageRecipient;
import ch.astorm.jotlmsg.OutlookMessageRecipient.Type;
import de.fricofresh.outlookspitter.CreateSplittedFilesParameter;

public class OutlookSplitterProcessorUtil {
	
	private static Logger log = LogManager.getLogger(OutlookSplitterProcessorUtil.class);
	
	public static List<OutlookMessageRecipient> receiveOutlookRecipients(String emailAdresses, Type type) {
		
		List<String> splitAdresses = OutlookMessageExtended.splitAdresses(emailAdresses);
		List<OutlookMessageRecipient> toOutlookRecipientsList = splitAdresses.stream().map(e -> new OutlookMessageRecipient(type, e)).collect(Collectors.toList());
		return toOutlookRecipientsList;
	}
	
	public static List<Path> createSplittedFiles(CreateSplittedFilesParameter parameterObject) {
		
		List<Path> result = new ArrayList<>();
		
		OutlookMessageExtended tempMessage = copyCloneMessage(parameterObject);
		
		for (int i = 0; i < parameterObject.getRecipientsToSplit().size(); i++) {
			OutlookMessageRecipient recipient = parameterObject.getRecipientsToSplit().get(i);
			if (i != 0 && i % parameterObject.getSplit() == 0) {
				tempMessage = writeTempMessage(parameterObject, result, tempMessage);
			}
			tempMessage.addRecipient(recipient);
		}
		writeTempMessage(parameterObject, result, tempMessage);
		
		return result;
	}
	
	public static OutlookMessageExtended writeTempMessage(CreateSplittedFilesParameter parameterObject, List<Path> result, OutlookMessageExtended tempMessage) {
		
		try {
			Path outputFile = MailSplitterUtil.getOutputFile(parameterObject, result.size());
			switch (parameterObject.getMailGenMehtod()) {
				case POICOPY:
				case POICLONE:
					tempMessage.writeTo(outputFile.toFile());
					break;
				case POIADVANCEDCOPY:
				case POIADVANCEDCLONE:
				default:
					tempMessage.writeToEditiable(outputFile.toFile());
					break;
			}
			result.add(outputFile);
		}
		catch (IOException | ChunkNotFoundException e) {
			log.error(e, e);
		}
		tempMessage = copyCloneMessage(parameterObject);
		return tempMessage;
	}
	
	public static OutlookMessageExtended copyCloneMessage(CreateSplittedFilesParameter parameterObject) {
		
		OutlookMessageExtended baseMessage = new OutlookMessageExtended(parameterObject.getEmailMessage());
		OutlookMessageExtended tempMessage;
		switch (parameterObject.getMailGenMehtod()) {
			case POICOPY:
			case POIADVANCEDCOPY:
				tempMessage = (OutlookMessageExtended) baseMessage.cloneCopy();
				break;
			case POICLONE:
			case POIADVANCEDCLONE:
			default:
				tempMessage = (OutlookMessageExtended) baseMessage.clone();
		}
		return tempMessage;
	}
	
	public static List<String> extractMailsFromFile(File file) {
		
		List<String> result = new ArrayList<>();
		try (InputStream inputStream = new FileInputStream(file); BufferedReader br = new BufferedReader(new InputStreamReader(inputStream))) {
			String line;
			while ((line = br.readLine()) != null) {
				result.add(line);
			}
		}
		catch (FileNotFoundException e) {
			log.error(e, e);
		}
		catch (IOException e) {
			log.error(e, e);
		}
		return result;
	}
	
}
