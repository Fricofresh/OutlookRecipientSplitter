# Outlook Recipient Splitter

A Dart CLI that mirrors the Java application's batch-splitting workflow for Outlook mail templates.

What it does
- Reads an input `.msg` file.
- Parses TO / CC / BCC recipient lists.
- Splits recipients into multiple output message files based on a configured batch size.
- Writes one output `.msg` per batch.
- Stores a companion JSON metadata file for each generated message.
- Optionally opens generated files with the default OS handler.

CLI usage

```bash
dart run bin/outlook_recipient_splitter.dart \
  -i "C:/input/template.msg" \
  -t "first@example.com;second@example.com" \
  -c "cc1@example.com;cc2@example.com" \
  -b "bcc1@example.com" \
  -s 25 \
  -st \
  -o "C:/output/splits"
```

Supported arguments
- `-i`, `--msginput` : source `.msg` file
- `-t`, `--to` : TO recipients
- `-c`, `--cc` : CC recipients
- `-b`, `--bcc` : BCC recipients
- `-st`, `--splitTo` : move TO recipients into the split list
- `-sc`, `--splitCC` : move CC recipients into the split list
- `-sb`, `--splitBCC` : move BCC recipients into the split list
- `-s`, `--split` : maximum recipients per generated file
- `-p`, `--prefix` : file-name prefix
- `-su`, `--suffix` : file-name suffix (default `.msg`)
- `-o`, `--outputdir` : output directory
- `-oc`, `--openCreated` : open generated files after creation
- `-mgm`, `--mailGenMethod` : generator method name (`POI`, `POIADVANCEDCOPY`, ...)
- `-h`, `--help` : show help

The Java implementation this project mirrors lives under:
`java/src/de/fricofresh/outlookspitter/`
