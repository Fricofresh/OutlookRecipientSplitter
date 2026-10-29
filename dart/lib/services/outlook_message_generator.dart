import 'dart:convert';
import 'dart:io';

import 'package:path/path.dart' as p;

import '../models/mail_gen_method.dart';
import '../models/outlook_recipient.dart';
import '../models/create_split_files_parameters.dart';

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

    outputDirectory.createSync(recursive: true);

    final outputFile = File(p.join(outputDirectory.path, outputFileName));
    outputFile.writeAsBytesSync(templateFile.readAsBytesSync(), flush: true);

    final metadata = {
      'templateFile': templateFile.path,
      'createdAtUtc': DateTime.now().toUtc().toIso8601String(),
      'mailGenMethod': mailGenMethod.name,
      'recipients': recipients.map((recipient) => recipient.toJson()).toList(),
    };

    final metadataFile = File('${p.withoutExtension(outputFile.path)}.json');
    metadataFile.writeAsStringSync(const JsonEncoder.withIndent('  ').convert(metadata));

    return outputFile;
  }
}
