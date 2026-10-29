import 'dart:convert';
import 'dart:io';

import 'package:path/path.dart' as p;

import '../models/mail_gen_method.dart';
import '../models/outlook_recipient.dart';
import '../models/create_split_files_parameters.dart';

class OutlookSplitterProcessor {
  static List<File> createSplitFiles(CreateSplitFilesParameters parameters) {
    final sourceFile = File(parameters.msgInputPath);
    if (!sourceFile.existsSync()) {
      throw FileSystemException('Input message file not found', parameters.msgInputPath);
    }

    final outputDirectory = _resolveOutputDirectory(parameters.outputDir);
    final recipientsForNonSplit = List<OutlookRecipient>.from(parameters.recipients);
    final splitRecipients = List<OutlookRecipient>.from(parameters.recipientsToSplit);

    if (splitRecipients.isEmpty) {
      final file = _createOutputFile(
        sourceFile: sourceFile,
        directory: outputDirectory,
        sourceName: p.basenameWithoutExtension(sourceFile.path),
        prefix: parameters.prefix ?? '',
        suffix: _normalizeSuffix(parameters.suffix ?? '.msg'),
        recipients: recipientsForNonSplit,
        sequenceNumber: 1,
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

      final file = _createOutputFile(
        sourceFile: sourceFile,
        directory: outputDirectory,
        sourceName: p.basenameWithoutExtension(sourceFile.path),
        prefix: parameters.prefix ?? '',
        suffix: _normalizeSuffix(parameters.suffix ?? '.msg'),
        recipients: recipientsForThisFile,
        sequenceNumber: files.length + 1,
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

  static File _createOutputFile({
    required File sourceFile,
    required Directory directory,
    required String sourceName,
    required String prefix,
    required String suffix,
    required List<OutlookRecipient> recipients,
    required int sequenceNumber,
  }) {
    final outputFile = File(p.join(directory.path, '${prefix}${sourceName}_$sequenceNumber$suffix'));
    outputFile.writeAsBytesSync(sourceFile.readAsBytesSync(), flush: true);

    final metadata = {
      'sourceFile': sourceFile.path,
      'createdAtUtc': DateTime.now().toUtc().toIso8601String(),
      'mailGenMethod': MailGenMethod.poi.name,
      'recipients': recipients.map((recipient) => recipient.toJson()).toList(),
    };

    final metadataFile = File('${p.withoutExtension(outputFile.path)}.json');
    metadataFile.writeAsStringSync(const JsonEncoder.withIndent('  ').convert(metadata));

    return outputFile;
  }

  static String _normalizeSuffix(String suffix) {
    final raw = suffix.trim();
    if (raw.isEmpty) {
      return '.msg';
    }
    return raw.startsWith('.') ? raw : '.$raw';
  }
}
