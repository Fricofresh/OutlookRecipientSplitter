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
	
	private final String guessedOutlookPath = "%ProgramFiles%\\Microsoft Office\\root\\";
	
	public static List<OutlookMessageRecipient> receiveOutlookRecipients(String emailAdresses, Type type) {
		
		List<String> splitAdresses = OutlookMessageExtended.splitAdresses(emailAdresses);
		List<OutlookMessageRecipient> toOutlookRecipientsList = splitAdresses.stream().map(e -> new OutlookMessageRecipient(type, e)).collect(Collectors.toList());
		return toOutlookRecipientsList;
	}
	
	public void openFiles(List<Path> files, Optional<String> outlookPath) throws IOException {
		
		Optional<Path> foundOutlookPath = Files.walk(Paths.get(guessedOutlookPath), 2).filter(e -> e.getFileName().equals("OUTLOOK.EXE")).findFirst();
		
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
		
		OutlookMessageExtended baseMessage = (OutlookMessageExtended) parameterObject.getEmailMessage();
		OutlookMessageExtended tempMessage = (OutlookMessageExtended) baseMessage.clone();
		
		for (int i = 0; i < parameterObject.getRecipientsToSplit().size(); i++) {
			// TODO set parameterObject.getRecipients as fixed for all messages
			OutlookMessageRecipient recipient = parameterObject.getRecipientsToSplit().get(i);
			try {
				if (i != 0 && i % parameterObject.getSplit() == 0) {
					tempMessage = writeTempMessage(parameterObject, result, baseMessage, tempMessage);
				}
				// TODO Loop to all Messages and export it after that
				tempMessage.addRecipient(recipient);
			}
			catch (IOException e) {
				e.printStackTrace();
			}
		}
		try {
			writeTempMessage(parameterObject, result, baseMessage, tempMessage);
		}
		catch (IOException e) {
			e.printStackTrace();
		}
		
		return result;
	}
	
	private static OutlookMessageExtended writeTempMessage(CreateSplittedFilesParameter parameterObject, List<Path> result, OutlookMessageExtended baseMessage, OutlookMessageExtended tempMessage) throws IOException {
		
		Path outputFile;
		if (!parameterObject.getOutputDir().isPresent())
			outputFile = Files.createTempFile(parameterObject.getPrefix().orElse("") + parameterObject.getEmailMessage().getSubject(), parameterObject.getSuffix().orElse(".msg"));
		else {
			outputFile = Files.createDirectories(new File(parameterObject.getOutputDir().get(), "").toPath());
			outputFile = Files.createFile(new File(parameterObject.getPrefix().orElse("") + parameterObject.getEmailMessage().getSubject() + parameterObject.getSuffix().orElse(".msg")).toPath());
		}
		result.add(outputFile);
		tempMessage.writeTo(outputFile.toFile());
		tempMessage = (OutlookMessageExtended) baseMessage.clone();
		return tempMessage;
	}
	
}
