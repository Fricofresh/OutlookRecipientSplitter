package de.fricofresh.outlookspitter.utils;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Optional;

import org.apache.poi.hsmf.exceptions.ChunkNotFoundException;

import de.fricofresh.outlookspitter.CreateSplittedFilesParameter;

public class MailSplitterUtil {
	
	private final static String guessedOutlookPath = "%ProgramFiles%\\Microsoft Office\\root\\";
	
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
		
		for (Path file : files)
			new ProcessBuilder().command("start", outlookPath.orElse(foundOutlookPath.get().toAbsolutePath().toString()), "/f", file.toAbsolutePath().toString()).start();
		// Alternative:
		// Runtime.getRuntime().exec("start",
		// outlookPath.orElse(foundOutlookPath.get().toAbsolutePath().toString()),
		// file.toAbsolutePath().toString()), new String[]
		// {pathToFile.toAbsolutePath().toString()});
	}
}
