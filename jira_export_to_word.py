import os, sys  # Standard Python Libraries
import xlwings as xw  # pip install xlwings
from docxtpl import DocxTemplate  # pip install docxtpl


# -- Documentation:
# python-docx-template: https://docxtpl.readthedocs.io/en/latest/

# Change path to current working directory
os.chdir(sys.path[0])

def main():
    ############ Make Your Configuration Here
    inputExcelName = "e-tablo.xlsx"#Excel name where the issues are
    inputSheet = "Sayfa1" #Sheet name where the issues are
    firstIssueCellIdentifier="A4" #title cell row Key | type | Summary | Priority etc with random order
    outputDocxName = "report.docx"
    templateDocxName = "template.docx"

    #Also match the column names below with jira export columns
    KeyCol="Key"
    SummaryCol=r"Bulgu Açıklaması"
    PriorityCol="Priority"
    ImpactCol="Etkisi"
    StatusCol="Status"
    ResolutionCol=r"Çözüm Önerisi"
    ReferencesCol=r"Referanslar"
    ##########################################

    wb = xw.Book(inputExcelName)
    shtGeneralReport = wb.sheets[inputSheet]

    # -- Get values from Excel
    context = shtGeneralReport.range(firstIssueCellIdentifier).expand().value

    excelRowCount=len(context)
    issueRowCount=excelRowCount-1

    templateDict={"issue":[]}

    oneissue = []
    for i in range(0,issueRowCount):

        oneissue.append({
                "Key":context[i+1][context[0].index(KeyCol)],
                "Summary": context[i+1][context[0].index(SummaryCol)],
                "Priority": context[i+1][context[0].index(PriorityCol)],
                "Impact": context[i+1][context[0].index(ImpactCol)],
                "Status": context[i+1][context[0].index(StatusCol)],
                "Resolution": context[i+1][context[0].index(ResolutionCol)],
                "References": context[i+1][context[0].index(ReferencesCol)]})
    templateDict["issue"]=oneissue


#    templateDict={
#        "issue":[
#            {
#            "Key":"1",
#            "Summary":"sum1",
#            "Priority":"p1",
#            "Etkisi":"e1",
#            "Status":"s1"
#            },{
#            "Key":"2",
#            "Summary": "sum2",
#            "Priority": "p2",
#            "Etkisi": "e2",
#            "Status": "s2"
#            }]}

# -- Read The Template
    doc = DocxTemplate(templateDocxName)

# -- Render & Save Word Document
    doc.render(templateDict)
    doc.save(outputDocxName)

if __name__ == "__main__":

    main()
