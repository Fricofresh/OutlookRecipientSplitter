import 'dart:io';

import '../models/mail_gen_method.dart';
import '../models/outlook_recipient.dart';

class OutlookMessageGenerator {
  static File buildFromTemplate({
    required File templateFile,
    required Directory outputDirectory,
    required String outputFileName,
    required List<OutlookRecipient> recipients,
    required MailGenMethod mailGenMethod,
  }) {
    if (!templateFile.existsSync()) {
      throw FileSystemException('Template .msg file not found', templateFile.path);
    }

    final outputFile = File('${outputDirectory.path}${Platform.pathSeparator}$outputFileName');
    outputDirectory.createSync(recursive: true);

    final bytes = templateFile.readAsBytesSync();
    outputFile.writeAsBytesSync(bytes, flush: true);

    final metadata = {
      'templateFile': templateFile.path,
      'createdAtUtc': DateTime.now().toUtc().toIso8601String(),
      'mailGenMethod': mailGenMethod.name,
      'recipients': recipients.map((recipient) => recipient.toJson()).toList(),
    };

    final metadataFile = File('${outputFile.path}.json');
    metadataFile.writeAsStringSync(
      const JsonEncoder.withIndent('  ').convert(metadata),
    );

    return outputFile;
  }
}
