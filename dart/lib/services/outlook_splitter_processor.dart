import 'dart:io';

import 'package:path/path.dart' as p;

import '../models/create_split_files_parameters.dart';
import '../models/mail_gen_method.dart';
import '../models/outlook_recipient.dart';
import 'outlook_message_generator.dart';

class OutlookSplitterProcessor {
  static List<File> createSplitFiles(CreateSplitFilesParameters parameters) {
    final sourceFile = File(parameters.msgInputPath);
    if (!sourceFile.existsSync()) {
      throw FileSystemException('Input message file not found', parameters.msgInputPath);
    }

    if (!parameters.msgInputPath.toLowerCase().endsWith('.msg')) {
      throw const FormatException('The source file must be a .msg template file.');
    }

    final outputDirectory = _resolveOutputDirectory(parameters.outputDir);
    final recipientsForNonSplit = List<OutlookRecipient>.from(parameters.recipients);
    final splitRecipients = List<OutlookRecipient>.from(parameters.recipientsToSplit);

    if (splitRecipients.isEmpty) {
      final outputName = _buildOutputFileName(
        baseName: p.basenameWithoutExtension(sourceFile.path),
        sequenceNumber: 1,
        prefix: parameters.prefix ?? '',
        suffix: parameters.suffix ?? '.msg',
      );

      final file = OutlookMessageGenerator.buildFromTemplate(
        templateFile: sourceFile,
        outputDirectory: outputDirectory,
        outputFileName: outputName,
        recipients: recipientsForNonSplit,
        mailGenMethod: parameters.mailGenMethod,
      );
      return <File>[file];
    }

    final files = <File>[];
    for (var index = 0; index < splitRecipients.length; index += parameters.split) {
      final end = (index + parameters.split < splitRecipients.length)
          ? index + parameters.split
          : splitRecipients.length;

      final batch = splitRecipients.sublist(index, end);
      final recipientsForThisFile = <OutlookRecipient>[]
        ..addAll(recipientsForNonSplit)
        ..addAll(batch);

      final outputName = _buildOutputFileName(
        baseName: p.basenameWithoutExtension(sourceFile.path),
        sequenceNumber: files.length + 1,
        prefix: parameters.prefix ?? '',
        suffix: parameters.suffix ?? '.msg',
      );

      final file = OutlookMessageGenerator.buildFromTemplate(
        templateFile: sourceFile,
        outputDirectory: outputDirectory,
        outputFileName: outputName,
        recipients: recipientsForThisFile,
        mailGenMethod: parameters.mailGenMethod,
      );

      files.add(file);
    }

    return files;
  }

  static Future<void> openFiles(List<File> files) async {
    for (final file in files) {
      if (!file.existsSync()) {
        continue;
      }

      if (Platform.isWindows) {
        await Process.start('cmd', ['/c', 'start', '', file.path]);
      } else if (Platform.isMacOS) {
        await Process.start('open', [file.path]);
      } else if (Platform.isLinux) {
        await Process.start('xdg-open', [file.path]);
      }
    }
  }

  static Directory _resolveOutputDirectory(String? outputDir) {
    final directory = outputDir == null || outputDir.trim().isEmpty
        ? Directory.systemTemp.createTempSync('outlook_splitter_')
        : Directory(outputDir);
    directory.createSync(recursive: true);
    return directory;
  }

  static String _buildOutputFileName({
    required String baseName,
    required int sequenceNumber,
    required String prefix,
    required String suffix,
  }) {
    final normalizedSuffix = suffix.trim();
    final finalSuffix = normalizedSuffix.isEmpty
        ? '.msg'
        : (normalizedSuffix.startsWith('.') ? normalizedSuffix : '.$normalizedSuffix');

    return '${prefix}${baseName}_$sequenceNumber$finalSuffix';
  }
}
