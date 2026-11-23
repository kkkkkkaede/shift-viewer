document.getElementById("checkButton").addEventListener(
    "click",
    async function() {

        const fileInput = document.getElementById("excelFile");
        const employeeName = document.getElementById("employeeName").value;

        // 名前が入力されているか確認
        if (employeeName.trim() === "") {
            alert("名前を入力してください");
            return;
        }

        // Excelファイルが選択されているか確認
        if (fileInput.files.length === 0) {
            alert("Excelファイルを選択してください");
            return;
        }

        const file = fileInput.files[0];

        // 前回の結果を削除
        document.getElementById("result").innerHTML = "";
        document.getElementById("total").textContent = "";

        const buffer = await file.arrayBuffer();

        const workbook = new ExcelJS.Workbook();

        await workbook.xlsx.load(buffer);
        let worksheet = workbook.worksheets[0];

        let allworktime=0;
        
        
        for(let sheetnum=0;sheetnum<workbook.worksheets.length;sheetnum++){
            let starttime=null;
            let starttimecell=null;
            let finishtime=0;
            let worktime=0;
            let worktimecellcount=0;
            worksheet = workbook.worksheets[sheetnum];
    // ==========================
    // 勤務日のシートか確認
    // ==========================
    // 勤務日のシートでなければ次のシートへ
    if (!isWorkSheet(worksheet.name)) {
        console.log("勤務表ではないためスキップ:", worksheet.name);
        continue;
    }

    console.log("勤務表:", worksheet.name);

    // ここから勤務時間の計算
            for(let i=1;i<100;i++){
                let cell1=worksheet.getCell("A"+i);
                const employeeName = document.getElementById("employeeName").value;
                const cellName = String(cell1.value ?? "").trim();
                const inputName = employeeName.trim();

                if (cellName === inputName) {
                    for(let j=1;j<40;j++){
                        let cell=worksheet.getCell(i,j);
                        const colors = [
                                        "FFFFFF00",
                                        "FFFFC000",
                                        "FFFF0000",
                                        "FF00B0F0",
                                        "FFFF99FF",
                                        "FF9933FF",
                                        "FF99FFCC"
                                    ];
                        let color = cell.fill?.fgColor?.argb;
                        if (colors.includes(color)||cell.fill?.fgColor?.theme === 0 &&cell.fill?.fgColor?.tint === -0.1499984740745262) {
                            if(starttime===null){
                                starttime=worksheet.getCell(3,j);
                                starttimecell=j;
                            }
                            worktime+=0.5;
                            worktimecellcount+=1;
                        }
                    }
                    if(worktime > 0 && starttime !== null){
                        finishtime=worksheet.getCell(3,starttimecell+worktimecellcount);
                        allworktime+=worktime;


                        starttimecell=Gettime(starttime);
                        finishtimecell=Gettime(finishtime);

                        console.log("勤務日時: ",worksheet.name," 勤務時間:",worktime," ");
                        console.log("worktimecellcount:",worktimecellcount);
                        console.log("starttimecell: ",starttimecell," ~ ","finishtime: ",finishtimecell);
                        const resultArea = document.getElementById("result");

                        const workElement = document.createElement("div");

                        workElement.className = "work";

                        workElement.innerHTML = `
                            <strong>${worksheet.name}</strong><br>
                            勤務時間：${worktime}時間<br>
                            ${starttimecell} ～ ${finishtimecell}
                        `;

    resultArea.appendChild(workElement);
                        break;
                    }
                }
            }
        }
        console.log("働く合計時間：",allworktime);
        document.getElementById("total").textContent =
        `合計勤務時間：${allworktime}時間`;
    }
)

function Gettime(timecell) {

    const time = timecell.value;

    // Dateの場合
    if (time instanceof Date) {
        const hour = String(time.getUTCHours()).padStart(2, "0");
        const minute = String(time.getUTCMinutes()).padStart(2, "0");
        const second = String(time.getUTCSeconds()).padStart(2, "0");

        return `${hour}:${minute}:${second}`;
    }

    // 文字列の場合
    if (typeof time === "string") {
        return time;
    }

    // 空セルの場合
    if (time === null || time === undefined) {
        return "";
    }

    // その他
    console.warn("時刻として認識できません:", time);
    return String(time);
}

function isWorkSheet(sheetName) {

    // 全角数字を半角数字に変換
    const name = sheetName.replace(/[０-９]/g, function(char) {
        return String.fromCharCode(char.charCodeAt(0) - 0xFEE0);
    });

    // 「○日」または「○月○日」に一致
    return /^(?:\d{1,2}月)?\d{1,2}日$/.test(name);
}


    /*color-argb
        清掃　　　　：FFFFFF00
        バール　　　：FFFFC000
        グランピング：FFFF0000
        フィットネス：FF00B0F0
        受付　　　　：FFFF99FF
        酵素　　　　：FF9933FF
        休憩        {テーマ色: 0
                        明暗: -0.1499984740745262}
        */




    //console.log("背景色：",cell.fill?.fgColor?.argb);