import datetime
from datetime import timedelta
from os import replace
from typing import Text
from zipfile import BadZipFile, error
from bs4.element import ResultSet

import pandas as pd
import pathlib
from pathlib import Path
import edinet
from edinet.xbrl_file import XBRLDir
import requests
from requests.models import REDIRECT_STATI
from urllib3.exceptions import InsecureRequestWarning, ProxySchemeUnknown, ResponseNotChunked, SSLError
#from requests.packages.urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

dateList = []
#取得期間の設定。スタートとエンド逆にしないよう注意
#本来の初日+1日
end_date = datetime.datetime.combine(datetime.date(2021,10,11), datetime.time(0, 0, 0)) 
#5年前のものから削除されていくので毎日更新する必要がある
start_date = datetime.datetime.combine(datetime.date(2016,11,2), datetime.time(0, 0, 0)) 

#df = pd.read_excel(Path.cwd().joinpath("StaticData").joinpath("上場企業リスト.xlsx"),sheet_name=0, header=0,usecols=[0])
df = pd.read_excel(Path.cwd().joinpath("StaticData").joinpath("東証12部上場企業リスト.xlsx"),sheet_name=0, header=0,usecols=[0])

listedCompany_secCode_List = list(df["SC"].dropna().apply(lambda x: str(int(x))))



#特定の企業のみ抽出したい場合はここで指定
listedCompany_secCode_List.clear()
listedCompany_secCode_List.append("25930")



print(listedCompany_secCode_List)
securities_report_docID_list = []
securities_report_edinetCode_list = []
error_edinetCode_dict = dict()

def daterange(_start, _end):
    for n in range((_end - _start).days):
        yield _start + timedelta(n)
for i in daterange(start_date.date(), end_date.date()):
    dateList.append(i)

print(dateList)

for day in dateList:

    url = "https://disclosure.edinet-fsa.go.jp/api/v1/documents.json"
    params = {"date": day, "type": 2}  
    res = requests.get(url, params=params).json()

    for num in range(len(res["results"])):
        sec_code= res["results"][num]["secCode"]
        #print(sec_code)
        docType_code= res["results"][num]["docTypeCode"]

        if docType_code =="120" and  sec_code in listedCompany_secCode_List:
            edinetCode = res["results"][num]["edinetCode"]
            docID = res["results"][num]["docID"]
            securities_report_edinetCode_list.append(edinetCode)
            securities_report_docID_list.append(docID)  
            print(len(securities_report_edinetCode_list))
            print(len(securities_report_docID_list))                
            print(str(day)+":"+str(sec_code))
                # if edinetCode in securities_report_edinetCode_list:
                #i = securities_report_edinetCode_list.index(edinetCode)
                #del securities_report_docID_list[i]
                #?securities_report_edinetCode_list.remove(edinetCode)? 
print(len(set(securities_report_docID_list)))
print(len(set(securities_report_edinetCode_list)))

edinetCode_df = pd.Series(securities_report_edinetCode_list)
with pd.ExcelWriter(Path.cwd().joinpath("StaticData").joinpath("東証12部上場企業リスト.xlsx"),mode="a",if_sheet_exists = "replace") as writer:
    edinetCode_df.to_excel(writer, sheet_name="有報を取得可能な企業")
print(len(securities_report_edinetCode_list))
print(len(securities_report_docID_list))         

#指標名＝シート名なので変な記号を入れるとエラー出るから注意
index_xbrlCode_dict = {
    
    #"提出会社の従業員数":["jpcrp_cor:NumberOfEmployees","CurrentYearInstant_NonConsolidatedMember"],
    #"提出会社の従業員の平均年齢_年":["jpcrp_cor:AverageAgeYearsInformationAboutReportingCompanyInformationAboutEmployees","CurrentYearInstant_NonConsolidatedMember"],
    #"提出会社の従業員の平均年齢_月":["jpcrp_cor:AverageAgeMonthsInformationAboutReportingCompanyInformationAboutEmployees","CurrentYearInstant_NonConsolidatedMember"],
    #"提出会社の従業員の平均勤続年数_年":["jpcrp_cor:AverageLengthOfServiceYearsInformationAboutReportingCompanyInformationAboutEmployees","CurrentYearInstant_NonConsolidatedMember"],
    #"提出会社の従業員の平均勤続年数_月":["jpcrp_cor:AverageLengthOfServiceMonthsInformationAboutReportingCompanyInformationAboutEmployees","CurrentYearInstant_NonConsolidatedMember"],
    #"提出会社の従業員の平均年間給与":["jpcrp_cor:AverageAnnualSalaryInformationAboutReportingCompanyInformationAboutEmployees","CurrentYearInstant_NonConsolidatedMember"],

    #"社外役員の状況":["jpcrp_cor:OutsideDirectorsAndOutsideCorporateAuditorsTextBlock","FilingDateInstant"]
    #"取締役の報酬等の総額(社外取締役を除く)":["jpcrp_cor:TotalAmountOfRemunerationEtcRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingOutsideDirectorsMember"],
    #"取締役の報酬等の総額(監査等委員及び社外取締役を除く)":["jpcrp_cor:TotalAmountOfRemunerationEtcRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingAuditAndSupervisoryCommitteeMembersAndOutsideDirectorsMember"]

    # "固定報酬(社外取締役を除く)":["jpcrp_cor:FixedRemunerationRemunerationByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingOutsideDirectorsMember"],
    # "固定報酬(監査等委員及び社外取締役を除く)":["jpcrp_cor:FixedRemunerationRemunerationByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingAuditAndSupervisoryCommitteeMembersAndOutsideDirectorsMember"],

    # "基本報酬(社外取締役を除く)":["jpcrp_cor:BaseRemunerationRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingOutsideDirectorsMember"],
    # "基本報酬(監査等委員及び社外取締役を除く)":["jpcrp_cor:BaseRemunerationRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingAuditAndSupervisoryCommitteeMembersAndOutsideDirectorsMember"],

    #"株式報酬(社外取締役を除く)":["jpcrp_cor:ShareAwardsRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingOutsideDirectorsMember"],
    #"株式報酬(監査等委員及び社外取締役を除く)":["jpcrp_cor:ShareAwardsRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingAuditAndSupervisoryCommitteeMembersAndOutsideDirectorsMember"]


    # "ストックオプション(社外取締役を除く)":["jpcrp_cor:ShareOptionRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingOutsideDirectorsMember"],
    # "ストックオプション(監査等委員及び社外取締役を除く)":["jpcrp_cor:ShareOptionRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingAuditAndSupervisoryCommitteeMembersAndOutsideDirectorsMember"],

    # "譲渡制限付株式報酬(社外取締役を除く)":["jpcrp_cor:RestrictedShareAwardsRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingOutsideDirectorsMember"],
    # "譲渡制限付株式報酬(監査等委員及び社外取締役を除く)":["jpcrp_cor:RestrictedShareAwardsRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingAuditAndSupervisoryCommitteeMembersAndOutsideDirectorsMember"],

    # "退職慰労金(社外取締役を除く)":["jpcrp_cor:RetirementBenefitsRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingOutsideDirectorsMember"],
    # "退職慰労金(監査等委員及び社外取締役を除く)":["jpcrp_cor:RetirementBenefitsRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingAuditAndSupervisoryCommitteeMembersAndOutsideDirectorsMember"],

    # "業績連動報酬(社外取締役を除く)":["jpcrp_cor:PerformanceBasedRemunerationRemunerationByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingOutsideDirectorsMember"],
    # "業績連動報酬(監査等委員及び社外取締役を除く)":["jpcrp_cor:PerformanceBasedRemunerationRemunerationByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingAuditAndSupervisoryCommitteeMembersAndOutsideDirectorsMember"],

    # "賞与(社外取締役を除く)":["jpcrp_cor:BonusRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingOutsideDirectorsMember"],
    # "賞与(監査等委員及び社外取締役を除く)":["jpcrp_cor:BonusRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingAuditAndSupervisoryCommitteeMembersAndOutsideDirectorsMember"],

    # "対象となる役員の員数(社外取締役を除く)":["jpcrp_cor:NumberOfDirectorsAndOtherOfficersRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingOutsideDirectorsMember"],
    # "対象となる役員の員数(監査等委員及び社外取締役を除く)":["jpcrp_cor:NumberOfDirectorsAndOtherOfficersRemunerationEtcByCategoryOfDirectorsAndOtherOfficers","CurrentYearDuration_DirectorsExcludingAuditAndSupervisoryCommitteeMembersAndOutsideDirectorsMember"]

    "企業名":["jpcrp_cor:CompanyNameCoverPage","FilingDateInstant"],
    #"従業員数":["jpcrp_cor:NumberOfEmployees","CurrentYearInstant"],
    #"研究開発費":["jpcrp_cor:ResearchAndDevelopmentExpensesIncludedInGeneralAndAdministrativeExpensesAndManufacturingCostForCurrentPeriod","CurrentYearDuration"],
    "売上高":["jpcrp_cor:NetSalesSummaryOfBusinessResults","CurrentYearDuration"],
    #"売上原価":["jppfs_cor:CostOfSales","CurrentYearDuration"],
    #"売上高2or営業収益":["jpcrp_cor:OperatingRevenue1SummaryOfBusinessResults","CurrentYearDuration"],
    #"販売費及び一般管理費":["jppfs_cor:SellingGeneralAndAdministrativeExpenses","CurrentYearDuration"],
    #"減価償却費":["jppfs_cor:DepreciationAndAmortizationOpeCF","CurrentYearDuration"],
    #???"事業収益":["jpcrp_cor:BusinessRevenueSummaryOfBusinessResults","CurrentYearDuration"]
    #"売上収益":["jpcrp_cor:RevenueIFRSSummaryOfBusinessResults","CurrentYearDuration"],
    "営業利益":["jppfs_cor:OperatingIncome","CurrentYearDuration"],
    #"受取利息":["jppfs_cor:InterestIncomeNOI","CurrentYearDuration"],
    #"受取配当金":["jppfs_cor:DividendsIncomeNOI","CurrentYearDuration"],
    #"支払利息":["jppfs_cor:InterestExpensesNOE","CurrentYearDuration"],
    #"現金及び預金":["jppfs_cor:CashAndDeposits","CurrentYearInstant"],
    #"棚卸資産":["jppfs_cor:Inventories","CurrentYearInstant"],
    #"原材料及び貯蔵品":["jppfs_cor:RawMaterialsAndSupplies","CurrentYearInstant"],
    #"原材料":["jppfs_cor:RawMaterials","CurrentYearInstant"],
    #"貯蔵品":["jppfs_cor:Supplies","CurrentYearInstant"],
    #"仕掛品":["jppfs_cor:WorkInProcess","CurrentYearInstant"],
    #"半製品":["jppfs_cor:SemiFinishedGoods","CurrentYearInstant"],
    #"商品及び製品":["jppfs_cor:MerchandiseAndFinishedGoods","CurrentYearInstant"],
    #"製品":["jppfs_cor:FinishedGoods","CurrentYearInstant"],
    #"商品":["jppfs_cor:Merchandise","CurrentYearInstant"],    
    #"売掛金":["jppfs_cor:AccountsReceivableTrade","CurrentYearInstant"],
    #"受取手形":["jppfs_cor:NotesReceivableTrade","CurrentYearInstant"],
    #"受取手形及び売掛金":["jppfs_cor:NotesAndAccountsReceivableTrade","CurrentYearInstant"],
    #"支払手形及び買掛金":["jppfs_cor:NotesAndAccountsPayableTrade","CurrentYearInstant"],
    #"買掛金":["jppfs_cor:AccountsPayableTrade","CurrentYearInstant"],
    #"支払手形":["jppfs_cor:NotesPayableTrade","CurrentYearInstant"],
    #"電子記録債権":["jppfs_cor:ElectronicallyRecordedMonetaryClaimsOperatingCA","CurrentYearInstant"],
    #"有価証券":["jppfs_cor:ShortTermInvestmentSecurities","CurrentYearInstant"],
    #!!!"投資有価証券":["jppfs_cor:InvestmentSecurities","CurrentYearInstant"],
    #"投資その他の資産合計":["jppfs_cor:InvestmentsAndOtherAssets","CurrentYearInstant"],
    #"短期借入金":["jppfs_cor:ShortTermLoansPayable","CurrentYearInstant"],
    #"一年以内返済の長期借入金":["jppfs_cor:CurrentPortionOfLongTermLoansPayable","CurrentYearInstant"],
    #"一年以内償還予定の社債":["jppfs_cor:CurrentPortionOfBonds","CurrentYearInstant"],
    #"流一年以内償還予定の転換社債":["jppfs_cor:CurrentPortionOfConvertibleBonds","CurrentYearInstant"],
    #"流動資産合計":["jppfs_cor:CurrentAssets","CurrentYearInstant"],
    #"固定資産合計":["jppfs_cor:NoncurrentAssets","CurrentYearInstant"],
    #"流動負債合計":["jppfs_cor:CurrentLiabilities","CurrentYearInstant"],
    #"固定負債合計":["jppfs_cor:NoncurrentLiabilities","CurrentYearInstant"],
    #"純資産":["jpcrp_cor:NetAssetsSummaryOfBusinessResults","CurrentYearInstant"],
    #"負債合計":["jppfs_cor:Liabilities","CurrentYearInstant"],
    #"負債純資産合計(資産合計,総資産)":["jppfs_cor:Assets","CurrentYearInstant"],
    #"新株予約権":["jppfs_cor:SubscriptionRightsToShares","CurrentYearInstant"],
    #"非支配株主持分":["jppfs_cor:NonControllingInterests","CurrentYearInstant"],
    #"親会社株主に帰属する当期純利益":["jpcrp_cor:ProfitLossAttributableToOwnersOfParentSummaryOfBusinessResults","CurrentYearDuration"],
    #"親会社の所有者に帰属する当期利益(IFRS)":["jpcrp_cor:ProfitLossAttributableToOwnersOfParentIFRSSummaryOfBusinessResults","CurrentYearDuration"],
    #"当期純利益":["jpcrp_cor:NetIncomeLossSummaryOfBusinessResults","CurrentYearDuration"] 
    #"税引前当期純利益":["jppfs_cor:IncomeBeforeIncomeTaxes","CurrentYearDuration"]
    #"流リース債務":["jppfs_cor:LeaseObligationsCL","CurrentYearInstant"],
    #"固長期借入金":["jppfs_cor:LongTermLoansPayable","CurrentYearInstant"],
    #"固社債":["jppfs_cor:BondsPayable","CurrentYearInstant"],
    #"固転換社債":["jppfs_cor:ConvertibleBonds","CurrentYearInstant"],
    #"固新株予約権付社債":["jppfs_cor:BondsWithSubscriptionRightsToSharesNCL","CurrentYearInstant"],
    #"固転換社債型新株予約権付社債":["jppfs_cor:ConvertibleBondTypeBondsWithSubscriptionRightsToShares","CurrentYearInstant"],
    #"固リース債務":["jppfs_cor:LeaseObligationsNCL","CurrentYearInstant"],
    #"投資活動によるキャッシュフロー":["jppfs_cor:NetCashProvidedByUsedInInvestmentActivities","CurrentYearDuration"],
    #"財務活動によるキャッシュフロー":["jppfs_cor:NetCashProvidedByUsedInFinancingActivities","CurrentYearDuration"],
    #"営業活動によるキャッシュフロー":["jppfs_cor:NetCashProvidedByUsedInOperatingActivities","CurrentYearDuration"]
    #"売上債権の増減額":["jppfs_cor:DecreaseIncreaseInNotesAndAccountsReceivableTradeOpeCF","CurrentYearDuration"],
    #"仕入債務の増減額":["jppfs_cor:IncreaseDecreaseInNotesAndAccountsPayableTradeOpeCF","CurrentYearDuration"],
    #"貸倒引当金の増減額":["jppfs_cor:IncreaseDecreaseInAllowanceForDoubtfulAccountsOpeCF","CurrentYearDuration"]
    
    #!!!"営業収益":"jpcrp_cor:OperatingRevenue2SummaryOfBusinessResults"
}
result = {}
#docIDを参照してZIP形式でダウンロード
DATA_ROOT = Path("D:")#Path.cwd().joinpath("data")
for (docID,edinetCode) in zip(securities_report_docID_list, securities_report_edinetCode_list):
    docPath_in_dataDir = DATA_ROOT.joinpath("raw").joinpath(docID)
    xbrl_path = Path()
    if docPath_in_dataDir.exists():
        xbrl_path = docPath_in_dataDir
        print("\n保存済:"+str(docPath_in_dataDir))
    else:
        xbrl_path = edinet.api.document.get_xbrl(docID, save_dir=DATA_ROOT.joinpath("raw"), expand_level="dir")
        print("\n未保存:"+str(xbrl_path))
    
    xbrl_dir = XBRLDir(xbrl_path)

    if edinetCode not in result.keys():      
        result[edinetCode] = {}       
        print("len(result):"+str(len(result)))

    #時系列なら"2021docID"のようにしないと。。。
    submit_date =""
    try:
        submit_date = xbrl_dir.xbrl.find("jpcrp_cor:FilingDateCoverPage").text    
        print("submitdate:"+ submit_date)
        submit_Year = datetime.datetime.strptime(submit_date, '%Y-%m-%d').year
        print(submit_Year)

        result[edinetCode][docID] = {}
        #"submit_Year"をキーとして各提出年をresult[docID]に格納していく
        result[edinetCode][docID]["submit_Year"] = submit_Year
        print(str(len(result)))
        print("[docID:"+docID+"]"+"[edinetCode:"+edinetCode+"]")
        for index_key in index_xbrlCode_dict.keys():
            try:
                print(index_key)
                all = xbrl_dir.xbrl.find_all(index_xbrlCode_dict[index_key][0])
                #これが空になっている
                #print(all)
                for foundIndex in all:
                    context_id = foundIndex._element["contextRef"]
                    #print(context_id)           
                    if context_id == index_xbrlCode_dict[index_key][1] and foundIndex.text != "":
                        index_text = foundIndex.text
                        result[edinetCode][docID][index_key] = index_text
                        print(result[edinetCode][docID][index_key])
                        break
                    elif context_id == (index_xbrlCode_dict[index_key][1]+"_NonConsolidatedMember") and foundIndex.text != "":
                        index_text = foundIndex.text
                        result[edinetCode][docID][index_key] = index_text
                        print(result[edinetCode][docID][index_key])
                        break         
            except(AttributeError,BadZipFile,FileNotFoundError,TypeError) as e:
                print(e)
                error_edinetCode_dict[edinetCode] = e
    except (AttributeError,BadZipFile,FileNotFoundError,TypeError) as e:
        print(e)
        error_edinetCode_dict[edinetCode] = e

#指標毎の処理
for index_key in index_xbrlCode_dict.keys():
    toExcel_row_company_item_lists = []
    toExcel_columns_list = ["EDINETコード"]
    #toExcel_columns_list.append(index_key)

    #期間内のすべての年度を入力
    toExcel_columns_list.extend([2021,2020,2019,2018,2017])


    result_edinetCode_list = result.keys()
    #企業ごとの処理
    for ecode in result_edinetCode_list:
        compamy_item_list = [ecode]
        for n in range(len(toExcel_columns_list)-1):
            compamy_item_list.append("")
        
        #時系列の処理
        for docID in result[ecode].keys():
            submit_year = int(result[ecode][docID]["submit_Year"])
            column_index = toExcel_columns_list.index(submit_year)
            compamy_item_list[column_index] = result[ecode][docID].get(index_key,"値なし")
            #print(ecode+"/"+docID+"/company_item_list:")
            #print(compamy_item_list)

        toExcel_row_company_item_lists.append(compamy_item_list)
        #print(ecode+"/toExcel_row_company_item_lists:")
        #print(toExcel_row_company_item_lists)

    result_df = pd.DataFrame(toExcel_row_company_item_lists,columns= toExcel_columns_list)
    with pd.ExcelWriter(Path.cwd().joinpath("StaticData").joinpath("あああ.xlsx"),mode="a",if_sheet_exists = "replace") as writer:
        result_df.to_excel(writer, sheet_name=index_key)

#print(result)
print(error_edinetCode_dict)
print(len(securities_report_edinetCode_list))
print(len(securities_report_docID_list))
print(len(error_edinetCode_dict))
print(len(result))
