enum RecipientType {
  to,
  cc,
  bcc,
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
        (value) => value.name == (json['type'] as String? ?? RecipientType.to.name),
        orElse: () => RecipientType.to,
      ),
      email: json['email'] as String? ?? '',
    );
  }
}
