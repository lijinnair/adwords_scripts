import com.google.appsscript.adwords._
import com.google.appsscript.mail._
import com.google.appsscript.base.Logger
import com.google.appsscript.spreadsheet._
import js.annotation.JSExport
import js.JSConverters._

/**
 * MCC-Script - Shopping SQRs for Search using Google Spreadsheet
 */
@JSExport
class findSearchQueries(spreadsheetUrl: String) {

  import AppsScarippedMain._

  // Script Description
  val scriptName = "Shopping SQRs for Search"
  val scriptVersion = "1.0"
  val scriptAuthor = "Heiko von Raussendorff"
  val scriptChangeLog = ""
  val scriptLicense = """
                      |Copyright 2015 crealytics GmbH
                      |
                      |Licensed under the Apache License, Version 2.0 (the "License");
                      |you may not use this file except in compliance with the License.
                      |You may obtain a copy of the License at
                      |
                      |    http://www.apache.org/licenses/LICENSE-2.0
                      |
                      |Unless required by applicable law or agreed to in writing, software
                      |distributed under the License is distributed on an "AS IS" BASIS,
                      |WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
                      |See the License for the specific language governing permissions and
                      |limitations under the License.""".stripMargin

  // Get config from setting tab
  val spreadsheet = AppsScarippedMain.spreadsheetByUrl(spreadsheetUrl)
  val currentAccountId = AdWordsApp.currentAccount().getCustomerId.replaceAll("-", "")
  val accountConfigs = AppsScarippedMain.getAccountConfigs(spreadsheet)
  val config = accountConfigs.find(_("accountId").toString.replaceAll("-", "") == currentAccountId).get
  val accountName = s"${config("name").toString} (${config("accountId").toString})"

  Logger.log(s"Start processing $accountName")

  // Get shopping campaign ids
  val shoppingCampaignIterator = AdWordsApp
    .shoppingCampaigns()
    .withCondition("Impressions > 100")
    .forDateRange("LAST_30_DAYS")
    .get();
  val shoppingCampaignIds = shoppingCampaignIterator.toList.map(c => c.asInstanceOf[js.Dynamic].getId().toString())
  val kpiCondition = buildKpiQuery(config)
  val initialDateRange = config("dateRange").toString.trim
  val dateRange = if (initialDateRange.isEmpty) "LAST_7_DAYS" else initialDateRange

  // Get search queries of shopping campaigns
  val query = s"""
              |SELECT Query, Impressions, Clicks, Cost, Ctr, ConvertedClicks, AverageCpc, CostPerConvertedClick,
              | ValuePerConvertedClick, ClickConversionRate, ConversionValue
              |FROM SEARCH_QUERY_PERFORMANCE_REPORT
              |WHERE
              | CampaignId IN [${shoppingCampaignIds.mkString(",")}]
              | $kpiCondition
              |DURING $dateRange""".stripMargin

  val report = AdWordsApp.report(query)

  // Read all rows and cast to dictionary because we want to use arbitrary fields
  val searchQueryPerformances = report.rows()
    .map(c => c.asInstanceOf[js.Dictionary[js.Any]])
    .toList
    .map { c => c("Query") = normalizeKeyword(c("Query").toString) ; c }
  val searchQuerySize = searchQueryPerformances.size
  val searchQueries = searchQueryPerformances.map(c => c("Query").toString).take(10000)

  // Get existing keywords from search campaigns
  val searchQueriesString = searchQueries.map(c => s"'$c'").mkString(",")

  val existingKeywordsSelector = AdWordsApp.keywords()
    .withCondition("Text IN [" + searchQueriesString + "]")
    .get()

  // strip all characters google is using for matchtype representation
  val existingKeywords = existingKeywordsSelector.map(c => normalizeKeyword(c.getText())).toSet

  // Identify new keywords
  val keywordProposals = searchQueryPerformances.filterNot(c => existingKeywords.contains(c("Query").toString))

  // Write result in spreadsheet
  val exportSheet = Option(spreadsheet.getSheetByName(config("name").toString))
    .getOrElse(spreadsheet.getSheetByName("Template").copyTo(spreadsheet).setName(config("name").toString))

  // Delete existing data
  if (exportSheet.getMaxRows > 8) {
    exportSheet.deleteRows(9, exportSheet.getMaxRows - 8)
  }

  //Add blank row
  exportSheet.appendRow(js.Array(" ", ""))

  // Warning: too many new proposals
  if (keywordProposals.size > 5000) {
    val warning = s"More than 5000 new keyword Proposals for $accountName! Run time of script may exceed the maximum execution time (approx. 30 minutes). Please adapt kpi settings accordingly!"
    Logger.log (warning)
    exportSheet.appendRow(js.Array(" ", "!!!", warning))
  }

  //
  // Write keyword proposals to sheet
  //
  // Define sorting
  val sortKey = config("kpi1").toString

  val keywordProposalsWithPerformance = keywordProposals
    .sortBy(-_(sortKey)
      .toString
      .replaceAllLiterally("%", "")
      .toDouble)

  for ((keywordProposal, index) <- keywordProposalsWithPerformance.zipWithIndex) {

    val perf = List("Impressions", "Clicks", "Ctr", "ConvertedClicks", "Cost", "AverageCpc", "CostPerConvertedClick", "ValuePerConvertedClick", "ClickConversionRate", "ConversionValue").map(keywordProposal)
    val newRow = js.Array("", (index + 1).toString, keywordProposal("Query")) ++ perf

    exportSheet.appendRow(newRow)
  }

  // Add account name and account id to sheet
  exportSheet.getRange("B2").setValue(accountName)
  exportSheet.getRange("D5").setValue(kpiCondition)

  // Create the date formatters
  val date = new js.Date()
  val lastRun = date.toLocaleDateString

  // Add last run value to sheet
  exportSheet.getRange("L5").setValue(lastRun)

  // Add count of keyword proposals to sheet
  val numberOfProposals = keywordProposalsWithPerformance.size.toString
  exportSheet.getRange("D5").setValue(numberOfProposals)

  // Prepare email
  val title = s"$numberOfProposals new keyword proposals found for the account ${config("name").toString} (${config("accountId").toString})"
  val emailBody = s"""Keyword Proposals from Shopping Search Queries
                 |-----------------------------------------------------------------------------------------
                 |
                 |$title
                 |
                 |Please check:
                 |$spreadsheetUrl
                 |
                 |-----------------------------------------------------------------------------------------
                 |crealytics adwords scripts - $scriptName (Version: $scriptVersion) - Author: $scriptAuthor
                 |-----------------------------------------------------------------------------------------
                 |
                 |$scriptLicense
                 |
                 |
                 |-----------------------------------------------------------------------------------------
                 |Change Log:
                 |$scriptChangeLog
                 |""".stripMargin
  sendEmail(emailBody, config("emails").toString, title, config("name").toString)

  Logger.log(title)
}

object AppsScarippedMain extends js.JSApp {

  def normalizeKeyword(kw: String): String = {
    kw.replaceAll( """[\[\]"']""", "")
  }

  def buildKpiQuery(config: Map[String, AnyRef]): String = {
    List("kpi1", "kpi2").map { field =>
      val condition = Array(config(field).toString.trim, config(field + "comp").toString.trim, config(field + "value").toString.trim)

      // Remove all incomplete KPI settings
      if (condition.forall(c => !c.isEmpty)) {
        " and " + condition.mkString(" ")
      } else ""
    }.mkString(" ")
  }

  def spreadsheetByUrl(spreadsheetUrl: String): Spreadsheet = {
    SpreadsheetApp.openByUrl(spreadsheetUrl);
  }

  def getAccountConfigs(spreadsheet: Spreadsheet): Seq[Map[String, AnyRef]] = {

    // Read data from Spreadsheet
    val data = spreadsheet.getRangeByName("data").getValues()
    val headers = data.head.map(_.toString)

    // Remove all rows, which are not activated or don't have values for account id or account name
    data.tail.map(row => headers.zip(row).toMap)
      .filterNot(_("accountId").toString.isEmpty)
      .filterNot(_("name").toString.isEmpty)
      .filter(_("activated").toString == "yes")
  }

  def sendEmail(body: String, emails: String, message: String, account: String): Unit = {

    if (!emails.isEmpty) {
      MailApp.sendEmail(emails, message, body)
    } else {
      Logger.log(s"Script can not send email because email field in the setting tab for the account $account is empty.")
    }

  }

  def main(): Unit = {}

  @JSExport
  def processAllAccounts(spreadsheetUrl: String): Unit = {
    val spreadsheet = spreadsheetByUrl(spreadsheetUrl)
    val accountConfigs = getAccountConfigs(spreadsheet)
    val ids = accountConfigs.map(config => config("accountId").toString.replaceAll("-", "")).toJSArray
    val accountsQuery = "ManagerCustomerId IN [" + ids.mkString(",") + "]"

    // Execute all accounts in parallel
    MccApp.accounts().withCondition(accountsQuery).asInstanceOf[js.Dynamic].executeInParallel("findSearchQueries", null, spreadsheetUrl)
  }
}
