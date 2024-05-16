package de.fricofresh.outlookspitter.utils;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hsmf.exceptions.ChunkNotFoundException;

import ch.astorm.jotlmsg.OutlookMessageRecipient;
import ch.astorm.jotlmsg.OutlookMessageRecipient.Type;
import de.fricofresh.outlookspitter.CreateSplittedFilesParameter;

public class MailSplitterUtil {
	
	private final static String guessedOutlookPath = "C:\\Program Files\\Microsoft Office\\root\\";
	
	private static Logger log = LogManager.getLogger(MailSplitterUtil.class);
	
	// TODO Export createMessage and addRecipient into a Interface for JavaMailMessageUtil and OutlookSplitterProcessorUtil
	// public static List<Path> createSplittedFiles(CreateSplittedFilesParameter parameterObject) {
	//
	// List<Path> result = new ArrayList<>();
	//
	// Message tempMessage = createMessage(parameterObject);
	//
	// for (int i = 0; i < parameterObject.getRecipientsToSplit().size(); i++) {
	// OutlookMessageRecipient recipient = parameterObject.getRecipientsToSplit().get(i);
	// try {
	// if (i != 0 && i % parameterObject.getSplit() == 0) {
	// tempMessage = createMessage(parameterObject);
	// }
	// tempMessage.addRecipient(convertOutlookRecipientsToJakarta(recipient), recipient.getAddress());
	// }
	// catch (MessagingException e) {
	// log.error(e, e);
	// }
	// }
	// writeTempMessage(parameterObject, result, tempMessage);
	//
	// return result;
	// }
	
	public static Path getOutputFile(CreateSplittedFilesParameter parameterObject, int currendCount) throws IOException, ChunkNotFoundException {
		
		Path outputFile;
		String fileEnding;
		
		switch (parameterObject.getMailGenMehtod()) {
			case JAVAMAIL:
				fileEnding = ".eml";
				break;
			case POICOPY:
			case POIADVANCEDCOPY:
			case POICLONE:
			case POIADVANCEDCLONE:
			case POI:
			default:
				fileEnding = ".msg";
				break;
		}
		if (parameterObject.getOutputDir().isPresent()) {
			outputFile = Files.createDirectories(new File(parameterObject.getOutputDir().get(), "").toPath());
			outputFile = Files.createFile(
					Paths.get(outputFile.toString(), parameterObject.getPrefix().orElse("") + parameterObject.getEmailMessage().getSubject() + "_" + currendCount + parameterObject.getSuffix().orElse(fileEnding)));
		}
		else {
			outputFile = Files.createTempFile(parameterObject.getPrefix().orElse("") + parameterObject.getEmailMessage().getSubject(), parameterObject.getSuffix().orElse(fileEnding));
		}
		return outputFile;
	}
	
	public static void openFiles(List<Path> files, Optional<String> outlookPath) throws IOException {
		
		Optional<Path> foundOutlookPath = Files.walk(Paths.get(guessedOutlookPath), 2).filter(e -> e.getFileName().toString().equals("OUTLOOK.EXE")).findFirst();
		
		for (Path file : files) {
			String openParam = "/f";
			if (file.toString().endsWith("eml"))
				openParam = "/eml";
			
			new ProcessBuilder().command(outlookPath.orElse(foundOutlookPath.get().toAbsolutePath().toString()), openParam, file.toAbsolutePath().toString()).start();
		}
		// Alternative:
		// Runtime.getRuntime().exec("start",
		// outlookPath.orElse(foundOutlookPath.get().toAbsolutePath().toString()), openParam,
		// file.toAbsolutePath().toString()), new String[]
		// {pathToFile.toAbsolutePath().toString()});
	}
	
	public static List<String> extractMailsFromFile(File file) {
		
		List<String> result = new ArrayList<>();
		try (InputStream inputStream = new FileInputStream(file); BufferedReader br = new BufferedReader(new InputStreamReader(inputStream))) {
			String line;
			while ((line = br.readLine()) != null) {
				if (line.contains(";")) {
					OutlookMessageExtended.splitAdresses(line).stream().map(OutlookMessageExtended::extractEmail).forEach(result::add);
					continue;
				}
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
	
	public static List<OutlookMessageRecipient> getOutlookRecipientsList(String[] emailAdresses, Type type) {
		
		return getOutlookRecipientsList(Arrays.asList(emailAdresses), type);
	}
	
	public static List<OutlookMessageRecipient> getOutlookRecipientsList(List<String> emailAdresses, Type type) {
		
		List<OutlookMessageRecipient> result = new ArrayList<>();
		if (emailAdresses.isEmpty())
			return result;
		
		for (String email : emailAdresses) {
			result.addAll(OutlookSplitterProcessorUtil.receiveOutlookRecipients(email, type));
		}
		return result;
	}
}
