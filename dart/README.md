# Outlook Recipient Splitter

A Dart CLI that mirrors the Java app's splitting workflow for Outlook message templates.

Overview
- Reads a source `.msg` template file.
- Parses TO / CC / BCC recipient lists.
- Splits recipients into multiple output files based on a configured batch size.
- Writes one `.msg` file per batch.
- Writes a matching JSON metadata file for each output.
- Optionally opens the created files in the system default application.

Usage

```bash
dart run bin/outlook_recipient_splitter.dart \
  -i "C:/input/template.msg" \
  -t "first@example.com;second@example.com" \
  -c "cc1@example.com;cc2@example.com" \
  -b "bcc1@example.com" \
  -s 25 \
  -st \
  -o "C:/temp/outlook-splits"
```

Supported flags
- `-i`, `--msginput`: input `.msg` file
- `-t`, `--to`: TO recipients
- `-c`, `--cc`: CC recipients
- `-b`, `--bcc`: BCC recipients
- `-st`, `--splitTo`: split TO recipients into separate batches
- `-sc`, `--splitCC`: split CC recipients into separate batches
- `-sb`, `--splitBCC`: split BCC recipients into separate batches
- `-s`, `--split`: maximum recipients per split file
- `-p`, `--prefix`: optional filename prefix
- `-su`, `--suffix`: optional filename suffix
- `-o`, `--outputdir`: output folder
- `-oc`, `--openCreated`: open generated files after creation
- `-mgm`, `--mailGenMethod`: generator metadata value
- `-h`, `--help`: show help

This rewrite targets the same core behavior as the Java project in:
`java/src/de/fricofresh/outlookspitter/`
