package de.fricofresh.outlookspitter.utils;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

import ch.astorm.jotlmsg.OutlookMessage;
import ch.astorm.jotlmsg.OutlookMessageRecipient;
import ch.astorm.jotlmsg.OutlookMessageRecipient.Type;
import de.fricofresh.outlookspitter.CreateSplittedFilesParameter;

public class OutlookSplitterProcessorUtil {
	
	private final String guessedOutlookPath = "%LOCALAPPDATA%\\Microsoft\\WindowsApps\\outlook.exe";
	
	public static List<OutlookMessageRecipient> receiveOutlookRecipients(String emailAdresses, Type type) {
		
		List<String> splitAdresses = OutlookMessageExtended.splitAdresses(emailAdresses);
		List<OutlookMessageRecipient> toOutlookRecipientsList = splitAdresses.stream()
				.map(e -> new OutlookMessageRecipient(type, e)).collect(Collectors.toList());
		return toOutlookRecipientsList;
	}
	
	public void openFiles(List<Path> files, Optional<String> outlookPath) throws IOException {
		
		for (Path file : files)
			new ProcessBuilder()
					.command("start", outlookPath.orElse(guessedOutlookPath), file.toAbsolutePath().toString()).start();
	}
	
	public static List<Path> createSplittedFiles(CreateSplittedFilesParameter parameterObject) {
		
		List<Path> result = new ArrayList<>();
		OutlookMessage tempMessage = parameterObject.getEmailMessage();
		for (int i = 0; i < parameterObject.getRecipientsToSplit().size(); i++) {
			// TODO set parameterObject.getRecipients as fixed for all messages
			OutlookMessageRecipient recipient = parameterObject.getRecipients().get(i);
			tempMessage.addRecipient(recipient);
			if (i % parameterObject.getSplit() == 0) {
				try {
					Path outputFile;
					if (!parameterObject.getOutputDir().isPresent())
						outputFile = Files.createTempFile(
								parameterObject.getPrefix().orElse("") + parameterObject.getEmailMessage().getSubject(),
								parameterObject.getSuffix().orElse(".msg"));
					else {
						outputFile = Files
								.createDirectories(new File(parameterObject.getOutputDir().get(), "").toPath());
						outputFile = Files.createFile(new File(
								parameterObject.getPrefix().orElse("") + parameterObject.getEmailMessage().getSubject()
										+ parameterObject.getSuffix().orElse(".msg")).toPath());
					}
					tempMessage.writeTo(outputFile.toFile());
					result.add(outputFile);
				}
				catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		
		return result;
	}
	
}
