param(
  [string] $subject = "subject",
  [string] $body = "body"
)

$PSEmailServer = "srv-kerio.gr.guord.local"
$mailto = "de@gr.guord.local"
$mailfrom = "srv-termx@gr.guord.local"
$subject_prefix = "ser-termx event:"

Send-MailMessage `
    -To $mailto `
    -from $mailfrom `
    -subject "$subject_prefix $subject" `
    -Body $body `
    -Encoding 'UTF8'
