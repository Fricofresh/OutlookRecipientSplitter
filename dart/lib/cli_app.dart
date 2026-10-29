import 'dart:io';

import 'package:args/args.dart';

import 'models/mail_gen_method.dart';
import 'models/create_split_files_parameters.dart';
import 'models/outlook_recipient.dart';
import 'services/outlook_splitter_processor.dart';

class CliApp {
  static Future<void> run(List<String> arguments) async {
    final parser = ArgParser();
    parser.addFlag('help', abbr: 'h', negatable: false, help: 'Show usage information.');
    parser.addOption('msginput', abbr: 'i', help: 'The path to the .msg file.');
    parser.addMultiOption('to', abbr: 't', help: 'TO recipient addresses.');
    parser.addMultiOption('cc', abbr: 'c', help: 'CC recipient addresses.');
    parser.addMultiOption('bcc', abbr: 'b', help: 'BCC recipient addresses.');
    parser.addFlag('splitTo', abbr: 'st', help: 'Split TO recipients into separate batches.');
    parser.addFlag('splitCC', abbr: 'sc', help: 'Split CC recipients into separate batches.');
    parser.addFlag('splitBCC', abbr: 'sb', help: 'Split BCC recipients into separate batches.');
    parser.addOption('split', abbr: 's', help: 'Maximum number of recipients per batch.');
    parser.addOption('prefix', abbr: 'p', help: 'Prefix for generated files.');
    parser.addOption('suffix', abbr: 'su', help: 'Suffix for generated files.');
    parser.addOption('outputdir', abbr: 'o', help: 'Directory for generated files.');
    parser.addFlag('openCreated', abbr: 'oc', help: 'Open generated files after creation.');
    parser.addOption('mailGenMethod', abbr: 'mgm', help: 'Mail generation method to report in metadata.');

    final results = parser.parse(arguments);

    if (results['help'] == true) {
      stdout.writeln('Outlook Recipient Splitter');
      stdout.writeln('');
      stdout.writeln(parser.usage);
      return;
    }

    final msgInput = results['msginput'] as String?;
    if (msgInput == null || msgInput.trim().isEmpty) {
      stderr.writeln('Missing required --msginput (-i).');
      stdout.writeln(parser.usage);
      exitCode = 64;
      return;
    }

    final splitRaw = results['split'] as String?;
    final splitValue = int.tryParse(splitRaw ?? '');
    if (splitValue == null || splitValue <= 0) {
      stderr.writeln('Missing or invalid --split (-s). Expected a positive integer.');
      stdout.writeln(parser.usage);
      exitCode = 64;
      return;
    }

    final inputFile = File(msgInput);
    if (!inputFile.existsSync()) {
      stderr.writeln('Input file does not exist: ${inputFile.path}');
      exitCode = 66;
      return;
    }

    try {
      final parameters = CreateSplitFilesParameters(
        msgInputPath: inputFile.path,
        recipients: _collectRecipients(results, 'to', RecipientType.to)
          ..addAll(_collectRecipients(results, 'cc', RecipientType.cc))
          ..addAll(_collectRecipients(results, 'bcc', RecipientType.bcc)),
        recipientsToSplit: _collectSplitRecipients(results),
        split: splitValue,
        prefix: results['prefix'] as String?,
        suffix: results['suffix'] as String?,
        outputDir: results['outputdir'] as String?,
        mailGenMethod: _parseMailGenMethod(results['mailGenMethod'] as String?),
      );

      final createdFiles = OutlookSplitterProcessor.createSplitFiles(parameters);
      for (final file in createdFiles) {
        stdout.writeln('Created: ${file.path}');
      }

      if (results['openCreated'] == true) {
        await OutlookSplitterProcessor.openFiles(createdFiles);
      }
    } on FileSystemException catch (error) {
      stderr.writeln('File system error: ${error.message}');
      exitCode = 1;
    } on FormatException catch (error) {
      stderr.writeln('Input parsing error: ${error.message}');
      exitCode = 65;
    } on Exception catch (error) {
      stderr.writeln('Unexpected error: $error');
      exitCode = 1;
    }
  }

  static List<OutlookRecipient> _collectRecipients(
    ArgResults results,
    String optionName,
    RecipientType type,
  ) {
    final rawValues = results[optionName] as List<String>? ?? const <String>[];
    final output = <OutlookRecipient>[];

    for (final rawValue in rawValues) {
      final addresses = rawValue.split(RegExp(r'[;,]'));
      for (final address in addresses) {
        final normalized = address.trim();
        if (normalized.isEmpty) {
          continue;
        }
        output.add(OutlookRecipient(type: type, email: normalized));
      }
    }
    return output;
  }

  static List<OutlookRecipient> _collectSplitRecipients(ArgResults results) {
    final output = <OutlookRecipient>[];

    if (results['splitTo'] == true) {
      output.addAll(_collectRecipients(results, 'to', RecipientType.to));
    }
    if (results['splitCC'] == true) {
      output.addAll(_collectRecipients(results, 'cc', RecipientType.cc));
    }
    if (results['splitBCC'] == true) {
      output.addAll(_collectRecipients(results, 'bcc', RecipientType.bcc));
    }

    return output;
  }

  static MailGenMethod _parseMailGenMethod(String? rawValue) {
    if (rawValue == null || rawValue.trim().isEmpty) {
      return MailGenMethod.poi;
    }

    final normalized = rawValue.trim().toUpperCase();
    return MailGenMethod.values.firstWhere(
      (method) => method.name.toUpperCase() == normalized,
      orElse: () => MailGenMethod.poi,
    );
  }
}
