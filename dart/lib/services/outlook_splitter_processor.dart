import 'mail_gen_method.dart';
import 'outlook_recipient.dart';

class CreateSplitFilesParameters {
  const CreateSplitFilesParameters({
    required this.msgInputPath,
    required this.recipients,
    required this.recipientsToSplit,
    required this.split,
    this.prefix,
    this.suffix,
    this.outputDir,
    this.mailGenMethod = MailGenMethod.poi,
  });

  final String msgInputPath;
  final List<OutlookRecipient> recipients;
  final List<OutlookRecipient> recipientsToSplit;
  final int split;
  final String? prefix;
  final String? suffix;
  final String? outputDir;
  final MailGenMethod mailGenMethod;
}
