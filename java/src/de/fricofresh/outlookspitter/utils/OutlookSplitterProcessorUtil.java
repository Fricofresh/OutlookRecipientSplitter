package de.fricofresh.outlookspitter.utils;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

import ch.astorm.jotlmsg.OutlookMessageRecipient;
import ch.astorm.jotlmsg.OutlookMessageRecipient.Type;
import de.fricofresh.outlookspitter.CreateSplittedFilesParameter;

public class OutlookSplitterProcessorUtil {
	
	private final static String guessedOutlookPath = "%ProgramFiles%\\Microsoft Office\\root\\";
	
	public static List<OutlookMessageRecipient> receiveOutlookRecipients(String emailAdresses, Type type) {
		
		List<String> splitAdresses = OutlookMessageExtended.splitAdresses(emailAdresses);
		List<OutlookMessageRecipient> toOutlookRecipientsList = splitAdresses.stream().map(e -> new OutlookMessageRecipient(type, e)).collect(Collectors.toList());
		return toOutlookRecipientsList;
	}
	
	public static void openFiles(List<Path> files, Optional<String> outlookPath) throws IOException {
		
		Optional<Path> foundOutlookPath = Files.walk(Paths.get(guessedOutlookPath), 2).filter(e -> e.getFileName().toString().equals("OUTLOOK.EXE")).findFirst();
		
		for (Path file : files)
			new ProcessBuilder().command("start", outlookPath.orElse(foundOutlookPath.get().toAbsolutePath().toString()), file.toAbsolutePath().toString()).start();
		// Alternative:
		// Runtime.getRuntime().exec("start",
		// outlookPath.orElse(foundOutlookPath.get().toAbsolutePath().toString()),
		// file.toAbsolutePath().toString()), new String[]
		// {pathToFile.toAbsolutePath().toString()});
	}
	
	public static List<Path> createSplittedFiles(CreateSplittedFilesParameter parameterObject) {
		
		List<Path> result = new ArrayList<>();
		
		OutlookMessageExtended tempMessage = copyCloneMessage(parameterObject);
		
		for (int i = 0; i < parameterObject.getRecipientsToSplit().size(); i++) {
			OutlookMessageRecipient recipient = parameterObject.getRecipientsToSplit().get(i);
			try {
				if (i != 0 && i % parameterObject.getSplit() == 0) {
					tempMessage = writeTempMessage(parameterObject, result, tempMessage);
				}
				tempMessage.addRecipient(recipient);
			}
			catch (IOException e) {
				e.printStackTrace();
			}
		}
		try {
			writeTempMessage(parameterObject, result, tempMessage);
		}
		catch (IOException e) {
			e.printStackTrace();
		}
		
		return result;
	}
	
	private static OutlookMessageExtended writeTempMessage(CreateSplittedFilesParameter parameterObject, List<Path> result, OutlookMessageExtended tempMessage) throws IOException {
		
		Path outputFile;
		if (!parameterObject.getOutputDir().isPresent())
			outputFile = Files.createTempFile(parameterObject.getPrefix().orElse("") + parameterObject.getEmailMessage().getSubject(), parameterObject.getSuffix().orElse(".msg"));
		else {
			outputFile = Files.createDirectories(new File(parameterObject.getOutputDir().get(), "").toPath());
			outputFile = Files
					.createFile(Path.of(outputFile.toString(), parameterObject.getPrefix().orElse("") + parameterObject.getEmailMessage().getSubject() + "_" + result.size() + parameterObject.getSuffix().orElse(".msg")));
		}
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
		tempMessage = copyCloneMessage(parameterObject);
		return tempMessage;
	}
	
	private static OutlookMessageExtended copyCloneMessage(CreateSplittedFilesParameter parameterObject) {
		
		OutlookMessageExtended baseMessage = (OutlookMessageExtended) parameterObject.getEmailMessage();
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
	
}
