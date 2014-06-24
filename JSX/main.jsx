/**
    Test of JSX
    @author Yang Yang
    @version Ver1.2
 **/
//#include 'extendables/extendables.jsx';
//================================================================================== INIT
// var PATH_BASE = "/c/Users/SuperSunSun/Dropbox/SigmaloveEDU/JSX/";
var PATH_BASE = "C:/Users/fly/Desktop/JSX/";
var META = {
    PATH_template: PATH_BASE + "template/test.indd",
    PATH_inputFolder: PATH_BASE + "input",
    PATH_outputFolder: PATH_BASE + "output"
};

//  计算运行时间
var OBJ_Timer = function () {
    var _time;
    return {
        setStart: function () {
            _time = (new Date()).getTime();
        },
        getDuration: function () {
            return _time != null ? (new Date()).getTime() - _time : 0;
        }
    }
}();
//================================================================================== PROCESS
OBJ_Timer.setStart();

var result = "";
main();
var _timeEnd = new Date().getTime();
result = OBJ_Timer.getDuration();
result;
//================================================================================== FUNCTIONS
function main() {

    //  1. 读取文件夹中文件
    //  2. 遍历XML文件
    //  3. 读取XML内容
    //  4. 将XML内容写入模板
    //  5. 生成ID文件
    //  6. ID文件导出成PDF

    //  1.
    var myFileList = openFolder(META.PATH_inputFolder);

    //  2.
    for (var i = 0; i < myFileList.length; i++) {
        var myFile = myFileList[i];
        if (myFile instanceof File && myFile.name.match(/\.xml$/i)) {

            //  3. 
            myFile.open("r");
            var xmlStr = myFile.read();
            myFile.close();


            var xmlRoot = new XML(xmlStr);
            app.open(new File(META.PATH_template));
            var myDocument = app.activeDocument;
            // var myDocument = app.documents.item(0);
            // clearTags (xmlStr);

            //  4,
            assignToDocument(myDocument, xmlRoot);

            //  6.
            if (true) { //  DEBUG 
                var myFolder = new Folder(META.PATH_outputFolder);
                if (!myFolder.exists) {
                    myFolder.create();
                }
                myDocument.exportFile(ExportFormat.pdfType, new File(META.PATH_outputFolder + "/" + myFile.name + ".pdf"));
                myDocument.close(SaveOptions.NO);
            }


        } else continue;
    }

}

function assignToDocument(myDocument, xmlData) {

    //  0. XML DTD配置与解释
    //  0. FOR LOOP
    //  1. 内容转码
    //  2. REQUIRE - HTML to INDESIGN
    //  3. 写入对应标签 

    //  init

    var myPages = myDocument.pages;
    // myDocument.textFrames.itemByName ("test").contents = "";
    //  Cover  - <frontPage>
        
    //var reflectProperties = myDocument.allPageItems;
    //alert_scroll("Object properties", reflectProperties.reflect.properties.sort());

    /*myDocument.textFrames.itemByName ("COVER_schoolName").contents = xmlData.child("frontPage")[0].child("schoolName")[0].toString();
    myDocument.textFrames.itemByName ("COVER_title").contents = "第三次月考";
    //myDocument.textFrames.itemByName ("COVER_studentName").contents = "张三";
    myDocument.textFrames.itemByName ("COVER_date").contents = "2012/12/21";*/

    //  Summary - <summary>
    // questioninfo-content-question-description

    var num = 0;
    //   Content - <sectionList><section><errQuestion>
    var xmlData_section = xmlData.child("sectionList")[0].child("section");
    var newPage = myPages.add();
    for (var index_section = 0; index_section < xmlData_section.length(); index_section++) {

        //  work for each <section>, read <section> data
        var xmlData_errQuestion = xmlData_section[index_section].child("errQuestion");
        myDocument.textFrames.item(-1).contents += "\n第" + (index_section + 1) + "节共" + xmlData_errQuestion.length() + "题：";


        for (var index_errQuestion = 0; index_errQuestion < xmlData_errQuestion.length(); index_errQuestion++) {
            //  work for each <errQuestion>
            var errQuestion = new errQuestionObj(xmlData_errQuestion[index_errQuestion]);
            //  myDocument.textFrames.itemByName ("test").contents += errQuestion.questionId+" ";
            myDocument.textFrames.item(-1).contents += errQuestion.questionId + " ";
            var newTextFrame = new Array();
            newTextFrame[num] = newPage.textFrames.add();

            //  转换html
            var returnList = parseHTML(errQuestion.description, myDocument, newPage, newTextFrame, newTextFrame[num]);
            var tableContent = returnList.newFrame;
            errQuestion.description = returnList.description;
            var imageFile = returnList.imageFile;

            if (tableContent instanceof Object) { // 含有表格
                newTextFrame[num] = newPage.textFrames.add();
                newTextFrame[num].contents = errQuestion.questionId + errQuestion.description.slice(0, errQuestion.description.indexOf("\n\n\n"));
                newTextFrame[num].geometricBounds = myGetBounds(myDocument, newPage, 20, 0);
                newTextFrame[num].geometricBounds = myGetBounds(myDocument, newPage, baseline, baseline + newTextFrame[num].texts[0].endBaseline);
                
                if (baseline + newTextFrame[num].texts[0].endBaseline > myDocument.documentPreferences.pageHeight){
                var newPage = myPages.add();
                newTextFrame[num].move(newPage);
                tableContent.move(newPage);
                newTextFrame[num].geometricBounds = myGetBounds(myDocument, newPage, 20, 0);
                newTextFrame[num].geometricBounds = myGetBounds(myDocument, newPage, 20, newTextFrame[num].texts[0].endBaseline);
                }

                tableContent.geometricBounds = myGetBounds(myDocument, newPage, 100, 0);
                tableContent.geometricBounds = myGetBounds(myDocument, newPage, newTextFrame[num].texts[0].endBaseline - 10, tableContent.paragraphs[0].endBaseline + 50);

                num++;
                newTextFrame[num] = newPage.textFrames.add();
                newTextFrame[num].contents = errQuestion.description.slice(errQuestion.description.indexOf("\n\n\n") + 4, errQuestion.description.length - 1);
                newTextFrame[num].geometricBounds = myGetBounds(myDocument, newPage, tableContent.paragraphs[0].endBaseline, 0);
                newTextFrame[num].geometricBounds = myGetBounds(myDocument, newPage, tableContent.paragraphs[0].endBaseline - 10, newTextFrame[num].texts[0].endBaseline);
                
            } else {
                newTextFrame[num].contents = errQuestion.questionId + errQuestion.description;
                newTextFrame[num].geometricBounds = myGetBounds(myDocument, newPage, 20, 0);
                if (num > 0) {
                    if (baseline + newTextFrame[num].texts[0].endBaseline < myDocument.documentPreferences.pageHeight) // 跨页
                        newTextFrame[num].geometricBounds = myGetBounds(myDocument, newPage, baseline, baseline + newTextFrame[num].texts[0].endBaseline);
                    else {
                        var newPage = myPages.add();
                        newTextFrame[num].move(newPage);
                        newTextFrame[num].geometricBounds = myGetBounds(myDocument, newPage, 20, 0);
                        newTextFrame[num].geometricBounds = myGetBounds(myDocument, newPage, 20, newTextFrame[num].texts[0].endBaseline);
                    }
                }
            }
            if (imageFile instanceof File) {
                try {
                    var imageGraphic = newPage.place(imageFile);
                    imageGraphic = imageGraphic[0];
                    var imageFrame = imageGraphic.parent;
                    
                    imageFrame.fit(FitOptions.PROPORTIONALLY);
                    imageFrame.move([newTextFrame[num].geometricBounds[1] + 75,newTextFrame[num].geometricBounds[0]]);
                   
                    
                } catch (e) {
                    alert(e);
                }
            }
            var baseline = newTextFrame[num].texts[0].endBaseline;
            num++;

            try {
                var myFont = app.fonts.item("宋体");
                newTextFrame[num - 1].paragraphs.item(0).appliedFont = myFont;
                // var myStyle = myDocument.paragraphStyles.item(1);
                // newTextFrame.paragraphs.item(0).appliedParagraphStyle = myStyle;
            } catch (e) {
                alert(e);
            }
            /* 跨页BUG, Ctrl+Alt+P 打开菜单勾选"对页"选项。
                 if(index_errQuestion % 2 == 0) {
                    newTextFrame.move([
                        myBounds[1] + myDocument.documentPreferences.pageWidth, 
                        myBounds[0] ]);                        
                 }
                  */
        }
    }
    myDocument.textFrames.item(-1).contents += "\n" + OBJ_Timer.getDuration() + "ms";
}


/**
    @param path : absolute path of a folder
    @return File object arrary or NULL if folder doesn't exist
**/
function openFolder(path) {
    var myFolder = new Folder(path);
    if (!myFolder.exists) return null;

    var tableContent = [];
    var myFileList = myFolder.getFiles();
    for (var i = 0; i < myFileList.length; i++) {
        var myFile = myFileList[i];

        // if (myFile instanceof File && myFile.name.match(/\.indd$/i)) {
        // app.open(myFile);
        //}

        tableContent.push(myFile);
    }
    return tableContent;
}


function myGetBounds(myDocument, myPage, top, bottom) {
    var myPageWidth = myDocument.documentPreferences.pageWidth;
    var myPageHeight = myDocument.documentPreferences.pageHeight;
    if (myPage.side == PageSideOptions.leftHand) {
        var myX2 = myPage.marginPreferences.left;
        var myX1 = myPage.marginPreferences.right;
    } else {
        var myX1 = myPage.marginPreferences.left;
        var myX2 = myPage.marginPreferences.right;
    }
    var myY1 = myPage.marginPreferences.top + top;
    var myX2 = myPageWidth - myX2;
    if (bottom == 0)
        var myY2 = myPageHeight - myPage.marginPreferences.bottom;
    else
        var myY2 = bottom;
    if (myPage.documentOffset % 2) // 左右留白
        return [myY1, myX1, myY2, myX2 - 30];
    else return [myY1, myX1 + 210, myY2, myX2 + 183];
}

/**
    Error Question Object
**/
function errQuestionObj(_XMLNode) {
    var XMLNode = _XMLNode;

    this.questionId = _getItemInNode(XMLNode, "questionId")
    this.qScore = _getItemInNode(XMLNode, "qScore");
    this.qIndex = _getItemInNode(XMLNode, "qIndex");
    this.answer = _getItemInNode(XMLNode, "answer");
    this.explanation = _getItemInNode(XMLNode, "explanation");
    this.qType = _getItemInNode(XMLNode, "qType");
    this.studentAnswer = _getItemInNode(XMLNode, "studentAnswer");
    this.description = _getItemInNode(XMLNode, "description");

    function _getItemInNode(parentNode, childName) {
        return parentNode[0].child("question")[0].child(childName)[0].child("item")[0].toString();
    }

}

function parseHTML(description, myDocument, newPage, newTextFrame) {
    var table_start = "<table cellspacing=\"0\" cellpadding=";
    //</fragment>
    // 匹配html标签
    do {
        var numLt = description.indexOf("<");
        var numGt = description.indexOf(">");
        var numTableStart = description.indexOf(table_start);
        var numTableEnd = description.indexOf("</table>") + 8;
        var sliceHtml = description.slice(numLt, numGt + 1);
        sliceAll = description.slice(numTableStart, numTableEnd);

        if (sliceHtml.match("data-optionvalue=\"")) {
            description = description.replace(sliceHtml, "");
        }
        switch (sliceHtml) {
        case "<span class=\"optext\">":
        case "</span>":
        case "<table border=\"0\" width=\"100%\" class=\"select_table\">":
        case "<tr>":
        case "</tr>":
        case "<p>":
        case "</table>":
        case "</div>":
        case "<br style=\"page-break-before:always; clear:both\"/>":
            description = description.replace(sliceHtml, "");
            break;
        default:
            description = description.replace(sliceHtml, "\n");
        }


        // 插入图片
        if (sliceHtml.match("<img src=")) {
            var imageLeft = sliceHtml.indexOf("images/") + 7;
            var imageRight = sliceHtml.indexOf("png") + 3;
            var imageName = sliceHtml.slice(imageLeft, imageRight);
            var imageFile = File(META.PATH_inputFolder + "/" + imageName);
        }


        if (sliceAll) { // 创建表，sliceAll包含整个table的标签与内容
            var tableReplace = sliceAll;
            do {
                var numLts = sliceAll.indexOf("<");
                var numGts = sliceAll.indexOf(">");
                var sliceTable = sliceAll.slice(numLts, numGts + 1); // s3为tags内容
                var leftSelect = sliceAll.indexOf("<td class=\"qoption\" data-optionvalue=");
                var rightSelect = sliceAll.indexOf("</td>");
                var tableSelect = sliceAll.slice(leftSelect, rightSelect + 5);
                var rowCount = 0;
                var colCount = 0;

                if (sliceTable.match("<tr style=")) { // 消除tags
                    sliceAll = sliceAll.replace(sliceTable, "#");
                    rowCount++;
                } else if (sliceTable.match("<td ")) {
                    sliceAll = sliceAll.replace(sliceTable, "@");
                    if (sliceTable.match("<td rowspan=")) {
                        var numRowSpan = sliceTable.substring(13, 14); //<td rowspan="2" 
                        var numRow1 = rowCount;
                        var numCol1 = colCount;
                    }
                    colCount++;
                    if (sliceTable.match("<td colspan=")) {
                        var numColSpan = sliceTable.substring(13, 14); //<td colspan="4" 
                        var numRow2 = rowCount;
                        var numCol2 = colCount;
                    }
                } else if (tableSelect != "") { // 选择题选项移出table
                    sliceAll = sliceAll.replace(tableSelect, "");
                    description += tableSelect;
                } else {
                    sliceAll = sliceAll.replace(sliceTable, "");
                }
            } while (numGts > -1);

            var newFrame = newPage.textFrames.add();
            newFrame.geometricBounds = myGetBounds(myDocument, newPage, 0, 0);
            while (sliceAll.replace("#@", "#") != sliceAll) {
                sliceAll = sliceAll.replace("#@", "#");
            }
            while ((sliceAll.substring(sliceAll.length - 1, sliceAll.length)) == "#") { // 消除多余的行列
                sliceAll = sliceAll.slice(0, sliceAll.length - 1);
            }
            sliceAll = sliceAll.slice(1, sliceAll.length);
            newFrame.contents += sliceAll;
            myTable = newFrame.texts.item(0).convertToTable("@", "#");

            if (numRowSpan > 1) { // 划分RowSpan
                var cellNumber = 1;
                for (var rowSpan = 1; rowSpan < numRowSpan; rowSpan++) {
                    var cellLength = myTable.rows[rowSpan].cells.length;
                    for (cellNumber; cellNumber < cellLength; cellNumber++) {
                        myTable.rows[rowSpan].cells[cellLength - cellNumber].contents = myTable.rows[rowSpan].cells[cellLength - cellNumber - 1].contents;
                    }
                    myTable.rows[rowSpan].cells[numCol1].contents = "";
                    myTable.rows[numRow1].cells[numCol1].merge(myTable.rows[rowSpan].cells[numCol1]);
                }
            }
            if (numColSpan > 1) { // 划分ColSpan
                myTable.rows[numRow2].cells[numCol2].merge(myTable.rows[numRow2].cells[numColSpan - numCol2 + 1]);
            }
            var tableArea = "\n\n\n";
            description = description.replace(tableReplace, tableArea);
            newFrame.geometricBounds = myGetBounds(myDocument, newPage, 100, newFrame.paragraphs[0].endBaseline);
        }
    } while (numGt > -1);
    return {
        description: description,
        newFrame: newFrame,
        imageFile: imageFile
    };
}

function alert_scroll(title, input) { //查看对象
    if (input instanceof Array)
        input = input.join("\r");
    var w = new Window("dialog", title);
    var list = w.add("edittext", undefined, input, {
        multiline: true,
        scrolling: true
    });
    list.maximumSize.height = w.maximumSize.height - 100;
    list.minimumSize.width = 250;
    w.add("button", undefined, "Close", {
        name: "ok"
    });
    w.show();
    /*
    // object
    var reflectProperties = app.colorSettings;
    // display properties
    alert_scroll("Object properties", reflectProperties.reflect.properties.sort());
    */
}