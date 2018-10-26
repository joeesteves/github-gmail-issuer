interface Issue {
  title: string
  body: string
  state: string
  html_url: string
  number: string
}

let issuers

function retrieveIssuers() {
  Logger.log("RETRIEVING")
  return ContactsApp.getContactGroup('Issuers')
    .getContacts()
    .map((contact) => {
      return ({
        email: contact.getPrimaryEmail(),
        company: contact.getCompanies()[0].getCompanyName()
      })
  })
}

function initIssuers() {
  let cacheObj = CacheService.getScriptCache(),
    cacheIssuers = cacheObj.get("issuers")
  if(!cacheIssuers){
    cacheObj.put("issuers", JSON.stringify(retrieveIssuers()))
    cacheIssuers = cacheObj.get("issuers")
  }
  return JSON.parse(cacheIssuers)
}

function buildAddOn(e) {
  issuers = initIssuers()
  // Activate temporary Gmail add-on scopes.
  GmailApp.setCurrentMessageAccessToken(e.messageMetadata.accessToken)
  const messageId = e.messageMetadata.messageId,
    senderData = extractSenderData(messageId)
  Logger.log(senderData)
  if (!senderData) return [emptyCard()]
  if (!senderData.company) return [emptyCard()]

  const resp = accessProtectedResource(
    `https://api.github.com/repos/ponyesteves/help-desk/issues?labels=${
      senderData.company
    }`
  )
  const issues: Issue[] = JSON.parse(resp)

  if (issues.length > 0) return issues.map(issue => buildIssueCard(issue))
  return [emptyCard()]
}

function emptyCard() {
  return CardService.newCardBuilder()
    .setHeader(
      CardService.newCardHeader().setTitle('No recent threads from this sender')
    )
    .build()
}

function extractSenderData(messageId) {
  const mail = GmailApp.getMessageById(messageId),
    senderEmail = extractEmailAddress(mail.getFrom())
  return issuers.filter(issuer => issuer.email == senderEmail)[0]
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
      // .setImageUrl(
      //   'https://cdn.iconscout.com/icon/free/png-256/issue-4-433271.png'
      // )
      .setTitle(`#${issue.number} ${issue.title}`)
  )
  var section = CardService.newCardSection().setHeader(
    `<font color="#1257e0">Issue #${issue.number}</font>`
  )
  section.addWidget(CardService.newTextParagraph().setText(issue.body))

  const threadLink = CardService.newOpenLink()
    .setUrl(issue.html_url)
    .setOpenAs(CardService.OpenAs.FULL_SIZE)
  var button = CardService.newTextButton()
    .setText('Go to Issue')
    .setOpenLink(threadLink)
  section.addWidget(CardService.newButtonSet().addButton(button))

  var clearCacheBtn = CardService.newTextButton()
  .setText('Update Contacts')
  .setOnClickAction(CardService.newAction().setFunctionName("clearCache"))

  section.addWidget(CardService.newButtonSet().addButton(clearCacheBtn))

  card.addSection(section)
  return card.build()
}

function clearCache() {
  CacheService.getScriptCache().remove("issuers")
}
