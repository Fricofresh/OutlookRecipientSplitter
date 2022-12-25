package de.fricofresh.outlookspitter;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;

import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;

import ch.astorm.jotlmsg.OutlookMessageRecipient;
import ch.astorm.jotlmsg.OutlookMessageRecipient.Type;
import de.fricofresh.outlookspitter.utils.OutlookMessageExtended;
import de.fricofresh.outlookspitter.utils.OutlookSplitterProcessorUtil;

public class OutlookSplitterCli {
	
	public static void main(String[] args) throws ParseException {
		
		// TODO extract Options to method
		Options cliOptions = new Options();
		
		Option msgFileOption = Option.builder().argName("i").longOpt("msginput").required(true)
				.desc("The Path to the .msg file").build();
		
		Option emailToAdressesOption = Option.builder().argName("t").longOpt("to").hasArgs().required(true)
				.desc("The E-Mail adresses to send it").build();
		
		// TODO Check if bool is getting returned
		Option splitToAdressesOption = Option.builder().argName("ts").longOpt("splitTo").hasArg(false)
				.desc("TO E-Mail-Adresses should be splitted").build();
		
		Option emailCCAdressesOption = Option.builder().argName("c").longOpt("cc").hasArgs().required(true)
				.desc("The E-Mail adresses to split").build();
		Option splitCCAdressesOption = Option.builder().argName("ts").longOpt("splitTo").hasArg(false)
				.desc("CC E-Mail-Adresses should be splitted").build();
		
		Option emailBCCAdressesOption = Option.builder().argName("b").longOpt("bcc").hasArgs().required(true)
				.desc("The E-Mail adresses to split").build();
		Option splitBCCAdressesOption = Option.builder().argName("ts").longOpt("splitTo").hasArg(false)
				.desc("BCC E-Mail-Adresses should be splitted").build();
		// TODO Change desc to something more understandable
		Option splitValueOption = Option.builder().argName("s").longOpt("split").hasArg().required(true)
				.desc("Amount to split the E-Mail adresses into").type(Integer.TYPE).build();
		// Option openAfterFinishedOption =
		// Option.builder().argName("").longOpt("").desc("Open files when
		// finished").build();
		Option outputDirOption = Option.builder().argName("o").longOpt("outputdir").hasArg().required(true).desc("")
				.build();
		
		// TODO add other Option
		cliOptions.addOption(msgFileOption);
		cliOptions.addOption(emailToAdressesOption);
		cliOptions.addOption(splitValueOption);
		
		DefaultParser defaultParser = new DefaultParser();
		CommandLine cmd = defaultParser.parse(cliOptions, args);
		Path filePath = new File(cmd.getOptionValue(msgFileOption)).toPath();
		try {
			CreateSplittedFilesParameter cSFParameter = new CreateSplittedFilesParameter();
			OutlookMessageExtended outlookMessage = new OutlookMessageExtended(filePath.toFile());
			String emailAdresses = cmd.getOptionValue(emailToAdressesOption);
			int splitValue = (Integer) cmd.getParsedOptionValue(splitValueOption);
			Optional<String> outputDir = Optional.ofNullable(cmd.getOptionValue(outputDirOption));
			
			// TODO Check if splitXXAdressOption is present or true
			List<OutlookMessageRecipient> toOutlookRecipientsList = OutlookSplitterProcessorUtil
					.receiveOutlookRecipients(emailAdresses, Type.TO);
			List<OutlookMessageRecipient> ccOutlookRecipientsList = OutlookSplitterProcessorUtil
					.receiveOutlookRecipients(emailAdresses, Type.CC);
			List<OutlookMessageRecipient> bccOutlookRecipientsList = OutlookSplitterProcessorUtil
					.receiveOutlookRecipients(emailAdresses, Type.BCC);
			
			cSFParameter.setEmailMessage(outlookMessage);
			cSFParameter.setOutputDir(outputDir);
			cSFParameter.setRecipients(toOutlookRecipientsList);
			cSFParameter.setSplit(splitValue);
			// cSFParameter.setPrefix();
			// cSFParameter.setSuffix();
			
			OutlookSplitterProcessorUtil.createSplittedFiles(cSFParameter);
		}
		catch (IOException e) {
			System.err.println("A Error accured at reading the .msg file.");
			e.printStackTrace();
		}
		
	}
	
}
