package de.fricofresh.outlookspitter;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import ch.astorm.jotlmsg.OutlookMessageRecipient;
import ch.astorm.jotlmsg.OutlookMessageRecipient.Type;
import de.fricofresh.outlookspitter.utils.OutlookMessageExtended;
import de.fricofresh.outlookspitter.utils.OutlookSplitterProcessorUtil;

public class OutlookSplitterCli {
	
	static Logger log = LogManager.getLogger(OutlookSplitterCli.class);
	
	public static void main(String[] args) throws ParseException {
		
		log.info("Outlooksplitter start");
		
		createCLIOptions(args);
		
		log.info("Outlooksplitter end");
	}
	
	private static void createCLIOptions(String[] args) throws ParseException {
		
		Options cliOptions = new Options();
		
		Option msgFileOption = Option.builder().option("i").longOpt("msginput").hasArg().required(true).desc("The Path to the .msg file").build();
		
		Option emailToAdressesOption = Option.builder().option("t").longOpt("to").argName("E-Mail-Adresses").hasArgs().required(false).desc("The E-Mail adresses to send it").build();
		Option splitToAdressesOption = Option.builder().option("st").longOpt("splitTo").hasArg(false).desc("TO E-Mail-Adresses should be splitted").build();
		
		Option emailCCAdressesOption = Option.builder().option("c").longOpt("cc").argName("E-Mail-Adresses").hasArgs().required(false).desc("The E-Mail adresses to split").build();
		Option splitCCAdressesOption = Option.builder().option("sc").longOpt("splitCC").hasArg(false).desc("CC E-Mail-Adresses should be splitted").build();
		
		Option emailBCCAdressesOption = Option.builder().option("b").longOpt("bcc").argName("E-Mail-Adresses").hasArgs().required(false).desc("The E-Mail adresses to split").build();
		Option splitBCCAdressesOption = Option.builder().option("sb").longOpt("splitBCC").hasArg(false).desc("BCC E-Mail-Adresses should be splitted").build();
		
		Option splitValueOption = Option.builder().option("s").longOpt("split").hasArg().required(true).desc("The number of email addresses when to split").type(Integer.TYPE).build();
		// TODO Add optional arguments for setting the outlook.exe path
		Option openAfterFinishedOption = Option.builder().option("oc").longOpt("openCreated").optionalArg(true).desc("Open files when finished").build();
		
		Option prefixFileNameOption = Option.builder().option("p").longOpt("prefix").hasArg().required(false).optionalArg(true).desc("Add a Prefix to the files").build();
		Option suffixFileNameOption = Option.builder().option("su").longOpt("suffix").hasArg().required(false).optionalArg(true).desc("Add a Suffix to the files").build();
		Option outputDirOption = Option.builder().option("o").longOpt("outputdir").hasArg().required(false).desc("").build();
		Option mailGenMethodOption = Option.builder().option("mgm").longOpt("mailGenMethod").hasArg().required(false).converter(MailGenMethod::valueOf)
				.desc("Trying to create Messages with other Methods. Default Method is POI. Methots are:" + MailGenMethod.values()).build();
		
		cliOptions.addOption(msgFileOption);
		cliOptions.addOption(emailToAdressesOption);
		cliOptions.addOption(splitToAdressesOption);
		cliOptions.addOption(emailCCAdressesOption);
		cliOptions.addOption(splitCCAdressesOption);
		cliOptions.addOption(emailBCCAdressesOption);
		cliOptions.addOption(splitBCCAdressesOption);
		cliOptions.addOption(splitValueOption);
		cliOptions.addOption(openAfterFinishedOption);
		cliOptions.addOption(prefixFileNameOption);
		cliOptions.addOption(suffixFileNameOption);
		cliOptions.addOption(outputDirOption);
		cliOptions.addOption(mailGenMethodOption);
		
		if (checkHelpCommand(args))
			printHelp(cliOptions);
		
		DefaultParser defaultParser = DefaultParser.builder().setStripLeadingAndTrailingQuotes(true).build();
		CommandLine cmd = defaultParser.parse(cliOptions, args);
		try {
			Path filePath = new File(cmd.getOptionValue(msgFileOption)).toPath();
			CreateSplittedFilesParameter cSFParameter = new CreateSplittedFilesParameter();
			OutlookMessageExtended outlookMessage = new OutlookMessageExtended(filePath.toFile());
			String[] emailToAdresses = cmd.getOptionValues(emailToAdressesOption);
			String[] emailCCAdresses = cmd.getOptionValues(emailCCAdressesOption);
			String[] emailBCCAdresses = cmd.getOptionValues(emailBCCAdressesOption);
			int splitValue = Integer.valueOf(cmd.getOptionValue(splitValueOption));
			Optional<String> outputDir = Optional.ofNullable(cmd.getOptionValue(outputDirOption));
			MailGenMethod mailGenMethod = cmd.getParsedOptionValue(mailGenMethodOption, MailGenMethod.POI);
			
			List<OutlookMessageRecipient> toOutlookRecipientsList = getOutlookRecipientsList(emailToAdresses, Type.TO);
			List<OutlookMessageRecipient> ccOutlookRecipientsList = getOutlookRecipientsList(emailCCAdresses, Type.CC);
			List<OutlookMessageRecipient> bccOutlookRecipientsList = getOutlookRecipientsList(emailBCCAdresses, Type.BCC);
			
			List<OutlookMessageRecipient> recipientsForAll = new ArrayList<>();
			List<OutlookMessageRecipient> recipientsToSplit = new ArrayList<>();
			
			if (cmd.hasOption(splitToAdressesOption))
				recipientsToSplit.addAll(toOutlookRecipientsList);
			else
				recipientsForAll.addAll(toOutlookRecipientsList);
			
			if (cmd.hasOption(splitCCAdressesOption))
				recipientsToSplit.addAll(ccOutlookRecipientsList);
			else
				recipientsForAll.addAll(ccOutlookRecipientsList);
			
			if (cmd.hasOption(splitBCCAdressesOption))
				recipientsToSplit.addAll(bccOutlookRecipientsList);
			else
				recipientsForAll.addAll(bccOutlookRecipientsList);
			
			cSFParameter.setEmailMessage(outlookMessage);
			cSFParameter.setOutputDir(outputDir);
			cSFParameter.setRecipients(recipientsForAll);
			cSFParameter.setRecipientsToSplit(recipientsToSplit);
			cSFParameter.setSplit(splitValue);
			cSFParameter.setMailGenMehtod(mailGenMethod);
			cSFParameter.setPrefix(Optional.ofNullable(cmd.getOptionValue(prefixFileNameOption)));
			cSFParameter.setSuffix(Optional.ofNullable(cmd.getOptionValue(suffixFileNameOption)));
			
			List<Path> createSplittedFiles = OutlookSplitterProcessorUtil.createSplittedFiles(cSFParameter);
			
			if (cmd.hasOption(openAfterFinishedOption)) {
				OutlookSplitterProcessorUtil.openFiles(createSplittedFiles, Optional.empty());
			}
		}
		catch (IOException e) {
			log.error("A Error accured at reading the .msg file.", e);
		}
		catch (Exception e) {
			printHelp(cliOptions);
			log.error(e);
		}
	}
	
	private static List<OutlookMessageRecipient> getOutlookRecipientsList(String[] emailAdresses, Type type) {
		
		if (emailAdresses == null)
			return new ArrayList<>();
		
		List<OutlookMessageRecipient> result = new ArrayList<>();
		for (String email : emailAdresses) {
			result.addAll(OutlookSplitterProcessorUtil.receiveOutlookRecipients(email, type));
		}
		
		return result;
	}
	
	private static boolean checkHelpCommand(String[] args) {
		
		String longName = "help";
		String shortName = "h";
		
		return Arrays.stream(args).anyMatch(s -> s.replace("-", "").equals(longName) || s.replace("-", "").equals(shortName));
		
	}
	
	private static void printHelp(Options options) {
		
		HelpFormatter helpFormatter = new HelpFormatter();
		helpFormatter.printHelp("OutlookSplitterCli.jar", options);
		
	}
	
}
