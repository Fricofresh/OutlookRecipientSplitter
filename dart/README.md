name: outlook_recipient_splitter

description: A production-ready Dart CLI for splitting Outlook .msg template batches by recipient.
version: 1.0.0
publish_to: 'none'

environment:
  sdk: '^3.3.0'

dependencies:
  args: ^2.5.0
  path: ^1.9.0

# No extra packages are required for the real .msg generation path because the app
# intentionally works by cloning a real Outlook .msg template and writing metadata.
