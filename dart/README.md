# Outlook Recipient Splitter (Dart rewrite)

This branch contains a Dart rewrite of the original Java CLI that splits a single Outlook `.msg` template into multiple output files based on recipient batches.

Important note
- The original Java implementation depends on the proprietary Outlook / MAPI API stack (`jotlmsg`, Apache POI, and Windows Outlook integration).
- This Dart port keeps the same CLI shape and splitting logic, but it is implemented as a portable Dart CLI that copies the base `.msg` file into batch outputs and stores a JSON sidecar with the recipients used for each generated file.
- This makes the project portable and keeps the behavior understandable without needing the Windows-only Outlook stack.

Usage

```bash
dart run bin/outlook_recipient_splitter.dart \
  -i "C:/path/to/template.msg" \
  -t "first@example.com;second@example.com" \
  -c "cc@example.com" \
  -b "bcc@example.com" \
  -s 25 \
  -o "C:/temp/outlook-splits"
```

Supported flags
- `-i`, `--msginput`: input `.msg` path
- `-t`, `--to`: TO recipients
- `-c`, `--cc`: CC recipients
- `-b`, `--bcc`: BCC recipients
- `-st`, `--splitTo`: send TO recipients to the split list
- `-sc`, `--splitCC`: send CC recipients to the split list
- `-sb`, `--splitBCC`: send BCC recipients to the split list
- `-s`, `--split`: split size
- `-p`, `--prefix`: filename prefix
- `-su`, `--suffix`: filename suffix
- `-o`, `--outputdir`: target folder
- `-oc`, `--openCreated`: open generated files after creation
- `-mgm`, `--mailGenMethod`: generator strategy (defaults to `POI`)

The Java source that this rewrite follows is located under `java/src/de/fricofresh/outlookspitter/`.
