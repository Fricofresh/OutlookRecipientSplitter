import 'dart:convert';
import 'dart:io';

import 'package:path/path.dart' as p;

enum RecipientType { to, cc, bcc }

enum MailGenMethod {
  poi,
  poiAdvancedCopy,
  poiAdvancedClone,
  poiCopy,
  poiClone,
}

class OutlookRecipient {
  const OutlookRecipient({
    required this.type,
    required this.email,
  });

  final RecipientType type;
  final String email;

  Map<String, dynamic> toJson() => {
        'type': type.name,
        'email': email,
      };

  factory OutlookRecipient.fromJson(Map<String, dynamic> json) {
    return OutlookRecipient(
      type: RecipientType.values.firstWhere(
        (entry) => entry.name == json['type'],
        orElse: () => RecipientType.to,
      ),
      email: (json['email'] as String?) ?? '',
    );
  }
}

class CreateSplitFilesParameters {
  const CreateSplitFilesParameters({
    required this.emailMessagePath,
    required this.recipients,
    required this.recipientsToSplit,
    required this.split,
    this.prefix,
    this.suffix,
    this.outputDir,
    this.mailGenMethod = MailGenMethod.poi,
  });

  final String emailMessagePath;
  final List<OutlookRecipient> recipients;
  final List<OutlookRecipient> recipientsToSplit;
  final int split;
  final String? prefix;
  final String? suffix;
  final String? outputDir;
  final MailGenMethod mailGenMethod;
}

class OutlookSplitterProcessor {
  static List<File> createSplitFiles(CreateSplitFilesParameters parameters) {
    final sourceFile = File(parameters.emailMessagePath);
    if (!sourceFile.existsSync()) {
      throw FileSystemException('Source .msg file not found', parameters.emailMessagePath);
    }

    final allRecipients = <OutlookRecipient>[]
      ..addAll(parameters.recipients)
      ..addAll(parameters.recipientsToSplit);

    final targetDirectory = parameters.outputDir == null || parameters.outputDir!.isEmpty
        ? Directory.systemTemp.createTempSync('outlook_splitter_')
        : Directory(parameters.outputDir!);
    targetDirectory.createSync(recursive: true);

    final generatedFiles = <File>[];
    final filePrefix = parameters.prefix ?? '';
    final suffix = parameters.suffix ?? '.msg';

    if (parameters.recipientsToSplit.isEmpty) {
      final file = _createFileFromSource(
        sourceFile: sourceFile,
        directory: targetDirectory,
        fileName: '${filePrefix}outlook_split_0$suffix',
        recipients: allRecipients,
      );
      generatedFiles.add(file);
      return generatedFiles;
    }

    final batches = <List<OutlookRecipient>>[];
    for (var i = 0; i < parameters.recipientsToSplit.length; i += parameters.split) {
      final end = (i + parameters.split < parameters.recipientsToSplit.length)
          ? i + parameters.split
          : parameters.recipientsToSplit.length;
      batches.add(parameters.recipientsToSplit.sublist(i, end));
    }

    for (var index = 0; index < batches.length; index++) {
      final batch = batches[index];
      final fileName = '${filePrefix}outlook_split_${index + 1}$suffix';
      final file = _createFileFromSource(
        sourceFile: sourceFile,
        directory: targetDirectory,
        fileName: fileName,
        recipients: [...parameters.recipients, ...batch],
      );
      generatedFiles.add(file);
    }

    return generatedFiles;
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

  static File _createFileFromSource({
    required File sourceFile,
    required Directory directory,
    required String fileName,
    required List<OutlookRecipient> recipients,
  }) {
    final outputFile = File(p.join(directory.path, fileName));
    outputFile.parent.createSync(recursive: true);

    final outputBytes = sourceFile.readAsBytesSync();
    outputFile.writeAsBytesSync(outputBytes, flush: true);

    final metadataFile = File('${outputFile.path}.json');
    final metadata = {
      'sourceFile': sourceFile.path,
      'createdAtUtc': DateTime.now().toUtc().toIso8601String(),
      'mailGenMethod': 'POI',
      'recipients': recipients.map((recipient) => recipient.toJson()).toList(),
    };
    metadataFile.writeAsStringSync(const JsonEncoder.withIndent('  ').convert(metadata));

    return outputFile;
  }
}

class OutlookMessageExtended {
  OutlookMessageExtended({
    required this.subject,
    required this.body,
    this.from,
    this.recipients = const <OutlookRecipient>[],
  });

  final String subject;
  final String body;
  final String? from;
  final List<OutlookRecipient> recipients;

  Map<String, dynamic> toJson() => {
        'subject': subject,
        'body': body,
        'from': from,
        'recipients': recipients.map((recipient) => recipient.toJson()).toList(),
      };
}

class OutlookSplitterCli {
  static void run(List<String> args) {
    final parser = ArgParser();
    parser.addOption('msginput', abbr: 'i');
    parser.addOption('split', abbr: 's');
    parser.addMultiOption('to', abbr: 't');
    parser.addMultiOption('cc', abbr: 'c');
    parser.addMultiOption('bcc', abbr: 'b');
    parser.addFlag('splitTo', abbr: 'st');
    parser.addFlag('splitCC', abbr: 'sc');
    parser.addFlag('splitBCC', abbr: 'sb');
    parser.addOption('outputdir', abbr: 'o');
    parser.addOption('prefix', abbr: 'p');
    parser.addOption('suffix', abbr: 'su');

    final parsed = parser.parse(args);
    final output = OutlookSplitterProcessor.createSplitFiles(
      CreateSplitFilesParameters(
        emailMessagePath: parsed['msginput'] as String? ?? '',
        recipients: const <OutlookRecipient>[],
        recipientsToSplit: const <OutlookRecipient>[],
        split: int.tryParse(parsed['split'] as String? ?? '0') ?? 0,
        prefix: parsed['prefix'] as String?,
        suffix: parsed['suffix'] as String?,
        outputDir: parsed['outputdir'] as String?,
      ),
    );

    for (final file in output) {
      stdout.writeln(file.path);
    }
  }
}

class OutlookSplitterProcessorUtil {
  static List<OutlookRecipient> receiveOutlookRecipients(String addresses, RecipientType type) {
    final normalized = addresses.split(RegExp(r'[;,]'));
    return normalized
        .map((address) => address.trim())
        .where((address) => address.isNotEmpty)
        .map((address) => OutlookRecipient(type: type, email: address))
        .toList(growable: false);
  }
}

String _normalizeEmailAddress(String email) {
  return email.trim().replaceAll(RegExp(r'\s+'), '');
}

List<String> splitAddresses(String addresses) {
  return addresses
      .split(RegExp(r'[;,]'))
      .map(_normalizeEmailAddress)
      .where((value) => value.isNotEmpty)
      .toList(growable: false);
}

class CreateSplittedFilesParameter extends CreateSplitFilesParameters {
  const CreateSplittedFilesParameter({
    required super.emailMessagePath,
    required super.recipients,
    required super.recipientsToSplit,
    required super.split,
    super.prefix,
    super.suffix,
    super.outputDir,
    super.mailGenMethod,
  });
}

Map<String, dynamic> ensureJsonObject(Map<String, dynamic> source) => source;
