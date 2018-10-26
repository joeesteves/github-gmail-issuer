const MAX_THREADS = 5

interface Issue {
  title: string
  body: string
  state: string
  html_url: string
  number: string
}

function buildAddOn(e) {
  const issues: Issue[] = JSON.parse(
    accessProtectedResource(
      'https://api.github.com/repos/ponyesteves/help-desk/issues'
    )
  )

  // Activate temporary Gmail add-on scopes.
  const accessToken = e.messageMetadata.accessToken
  GmailApp.setCurrentMessageAccessToken(accessToken)

  const messageId = e.messageMetadata.messageId,
    senderData = extractSenderData(messageId)
  let cards = []

  if (issues.length > 0) {
    issues.forEach(issue => {
      cards = [...cards, buildIssueCard(issue)]
    })
  } else {
    cards = [emptyCard()]
  }

  return cards
}

function emptyCard() {
  CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader().setTitle('No recent threads from this sender')
    )
    .build()
}
/**
 *  This function builds a set of data about this sender's presence in your
 *  inbox.
 *
 *  @param {String} messageId The message ID of the open message.
 *  @return {Object} a collection of sender information to display in cards.
 */
function extractSenderData(messageId) {
  // Use the Gmail service to access information about this message.
  const mail = GmailApp.getMessageById(messageId),
    senderEmail = extractEmailAddress(mail.getFrom())
  return {
    email: senderEmail,
  }
}

function extractEmailAddress(sender) {
  var regex = /\<([^\@]+\@[^\>]+)\>/
  var email = sender // Default to using the whole string.
  var match = regex.exec(sender)
  if (match) {
    email = match[1]
  }
  return email
}

function buildIssueCard(issue: Issue) {
  const card = CardService.newCardBuilder().setHeader(

    CardService.newCardHeader()
    .setImageUrl("https://cdn.iconscout.com/icon/free/png-256/issue-4-433271.png")
    .setTitle(`#${issue.number} ${issue.title}`)
  )
  var section = CardService.newCardSection().setHeader(
    `<font color="#1257e0">Issue #${issue.number}</font>`
  )
  section.addWidget(CardService.newTextParagraph().setText(issue.body))
  // section.addWidget(
  //   CardService.newKeyValue()
  //     .setTopLabel('Sender')
  //     .setContent(senderEmail)
  // )
  // section.addWidget(
  //   CardService.newKeyValue()
  //     .setTopLabel('Number of messages')
  //     .setContent(threadData.count.toString())
  // )
  // section.addWidget(
  //   CardService.newKeyValue()
  //     .setTopLabel('Last updated')
  //     .setContent(threadData.lastDate.toString())
  // )

  const threadLink = CardService.newOpenLink()
    .setUrl(issue.html_url)
    .setOpenAs(CardService.OpenAs.FULL_SIZE)
  var button = CardService.newTextButton()
    .setText('Go to Issue')
    .setOpenLink(threadLink)
  section.addWidget(CardService.newButtonSet().addButton(button))

  card.addSection(section)
  return card.build()
}

/**
 *  Builds a card to display information about a recent thread from this sender.
 *
 *  @param {String} senderEmail The sender email.
 *  @param {Object} threadData Infomation about the thread to display.
 *  @return {Card} a card that displays thread information.
 */

function buildRecentThreadCard(senderEmail, threadData) {
  var card = CardService.newCardBuilder()
  card.setHeader(CardService.newCardHeader().setTitle(threadData.subject))
  var section = CardService.newCardSection().setHeader(
    '<font color="#1257e0">Recent thread</font>'
  )
  section.addWidget(CardService.newTextParagraph().setText(threadData.subject))
  section.addWidget(
    CardService.newKeyValue()
      .setTopLabel('Sender')
      .setContent(senderEmail)
  )
  section.addWidget(
    CardService.newKeyValue()
      .setTopLabel('Number of messages')
      .setContent(threadData.count.toString())
  )
  section.addWidget(
    CardService.newKeyValue()
      .setTopLabel('Last updated')
      .setContent(threadData.lastDate.toString())
  )

  var threadLink = CardService.newOpenLink()
    .setUrl(threadData.link)
    .setOpenAs(CardService.OpenAs.FULL_SIZE)
  var button = CardService.newTextButton()
    .setText('Open Issue')
    .setOpenLink(threadLink)
  section.addWidget(CardService.newButtonSet().addButton(button))

  card.addSection(section)
  return card.build()
}
