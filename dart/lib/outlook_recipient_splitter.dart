import 'dart:convert';
import 'dart:io';

import 'package:args/args.dart';
import 'package:path/path.dart' as p;

import '../lib/outlook_recipient_splitter.dart';

Future<void> main(List<String> arguments) async {
  final parser = ArgParser();
  parser.addFlag('help', abbr: 'h', negatable: false, help: 'Show this help screen.');
  parser.addOption(
    'msginput',
    abbr: 'i',
    help: 'The path to the .msg file.',
  );
  parser.addMultiOption(
    'to',
    abbr: 't',
    help: 'The TO E-Mail addresses to send it to.',
    valueHelp: 'E-Mail-Addresses',
  );
  parser.addFlag('splitTo', abbr: 'st', help: 'TO recipients should be split.');
  parser.addMultiOption(
    'cc',
    abbr: 'c',
    help: 'The CC E-Mail addresses to split.',
    valueHelp: 'E-Mail-Addresses',
  );
  parser.addFlag('splitCC', abbr: 'sc', help: 'CC recipients should be split.');
  parser.addMultiOption(
    'bcc',
    abbr: 'b',
    help: 'The BCC E-Mail addresses to split.',
    valueHelp: 'E-Mail-Addresses',
  );
  parser.addFlag('splitBCC', abbr: 'sb', help: 'BCC recipients should be split.');
  parser.addOption(
    'split',
    abbr: 's',
    help: 'The number of recipients to include in each split batch.',
  );
  parser.addFlag('openCreated', abbr: 'oc', help: 'Open files when finished.');
  parser.addOption('prefix', abbr: 'p', help: 'Add a file prefix.');
  parser.addOption('suffix', abbr: 'su', help: 'Add a file suffix.');
  parser.addOption('outputdir', abbr: 'o', help: 'Output directory for generated files.');
  parser.addOption(
    'mailGenMethod',
    abbr: 'mgm',
    help: 'Message generation method. Supported values: ${MailGenMethod.values.map((e) => e.name).join(', ')}',
  );

  final options = parser.parse(arguments);

  if (options['help'] as bool) {
    _printHelp(parser);
    return;
  }

  final msgInput = options['msginput'] as String?;
  if (msgInput == null || msgInput.isEmpty) {
    stderr.writeln('Missing required --msginput (-i).');
    _printHelp(parser);
    exitCode = 64;
    return;
  }

  final splitSizeRaw = options['split'] as String?;
  if (splitSizeRaw == null || int.tryParse(splitSizeRaw) == null) {
    stderr.writeln('Missing required --split (-s) value.');
    _printHelp(parser);
    exitCode = 64;
    return;
  }

  final splitSize = int.parse(splitSizeRaw);
  if (splitSize <= 0) {
    stderr.writeln('The split value must be greater than zero.');
    exitCode = 64;
    return;
  }

  final sourceFile = File(msgInput);
  if (!sourceFile.existsSync()) {
    stderr.writeln('Input file does not exist: $msgInput');
    exitCode = 66;
    return;
  }

  final outputDir = options['outputdir'] as String?;
  final mailGenMethod = _parseMailGenMethod(options['mailGenMethod'] as String?);

  final recipientsForAll = <OutlookRecipient>[];
  final recipientsToSplit = <OutlookRecipient>[];

  final toRecipients = _parseRecipients(options['to'] as List<String>?, RecipientType.to);
  final ccRecipients = _parseRecipients(options['cc'] as List<String>?, RecipientType.cc);
  final bccRecipients = _parseRecipients(options['bcc'] as List<String>?, RecipientType.bcc);

  if (options['splitTo'] as bool) {
    recipientsToSplit.addAll(toRecipients);
  } else {
    recipientsForAll.addAll(toRecipients);
  }

  if (options['splitCC'] as bool) {
    recipientsToSplit.addAll(ccRecipients);
  } else {
    recipientsForAll.addAll(ccRecipients);
  }

  if (options['splitBCC'] as bool) {
    recipientsToSplit.addAll(bccRecipients);
  } else {
    recipientsForAll.addAll(bccRecipients);
  }

  final parameters = CreateSplitFilesParameters(
    emailMessagePath: sourceFile.path,
    recipients: recipientsForAll,
    recipientsToSplit: recipientsToSplit,
    split: splitSize,
    prefix: options['prefix'] as String?,
    suffix: options['suffix'] as String?,
    outputDir: outputDir,
    mailGenMethod: mailGenMethod,
  );

  try {
    final createdFiles = OutlookSplitterProcessor.createSplitFiles(parameters);
    for (final createdFile in createdFiles) {
      stdout.writeln('Created: ${createdFile.path}');
    }

    if (options['openCreated'] as bool) {
      await OutlookSplitterProcessor.openFiles(createdFiles);
    }
  } on Exception catch (error) {
    stderr.writeln(error);
    exitCode = 1;
  }
}

void _printHelp(ArgParser parser) {
  stdout.writeln('Outlook Recipient Splitter (Dart port)');
  stdout.writeln('');
  stdout.writeln(parser.usage);
}

MailGenMethod _parseMailGenMethod(String? rawValue) {
  if (rawValue == null || rawValue.isEmpty) {
    return MailGenMethod.poi;
  }

  final value = rawValue.trim().toUpperCase();
  for (final method in MailGenMethod.values) {
    if (method.name.toUpperCase() == value) {
      return method;
    }
  }

  throw ArgumentError(
    'Unknown mail generation method: $rawValue. Supported values: ${MailGenMethod.values.map((m) => m.name).join(', ')}',
  );
}

List<OutlookRecipient> _parseRecipients(List<String>? values, RecipientType type) {
  if (values == null || values.isEmpty) {
    return const <OutlookRecipient>[];
  }

  final all = <OutlookRecipient>[];
  for (final entry in values) {
    final addresses = entry.split(RegExp(r'[;,]'));
    for (final address in addresses) {
      final trimmed = address.trim();
      if (trimmed.isEmpty) {
        continue;
      }
      all.add(OutlookRecipient(type: type, email: trimmed));
    }
  }

  return all;
}
