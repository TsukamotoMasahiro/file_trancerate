const fs = require("fs");
const XlsxPopulate = require('xlsx-populate');


  // 空のワークブックを作成する。
  XlsxPopulate.fromBlankAsync().then(workbook => {
    //メモリが限界の時は
    //export NODE_OPTIONS="--max-old-space-size=4096"
    
    
    function Main(){
      for (k=0;k<=20;++k){
        
        var x = buff.split("\n")[1].split(" ").map((n)=> n);
        //1行目の五行目に全体の行数が出力されたファイルのため、次のように定義
        var N = parseInt(x[4],10)
        //21行分のデータが欲しかったので、そのように定義
        var N1 = N+20
        var N2 = N-2
        var b = buff.split("\n")[N-1].split("	").map((n)=> n);
  
        var hako =["A1","B1","C1","D1","E1","F1","G1","H1","I1","J1","K1","L1","M1","N1","O1"
                ,"P1","Q1","R1","S1","T1","U1"] 
        workbook.sheet(j).cell(hako[k]).value(b[k]);
        for (i= N ; i <= N1 ;i++){
          var a = buff.split("\n")[i].split("	").map((n)=> Number(n));
          var p = i-N2
          var hakoo =[`A${p}`,`B${p}`,`C${p}`,`D${p}`,`E${p}`,`F${p}`
                  ,`G${p}`,`H${p}`,`I${p}`,`J${p}`,`K${p}`,`L${p}`,`M${p}`,`N${p}`,`O${p}`
                ,`P${p}`,`Q${p}`,`R${p}`,`S${p}`,`T${p}`,`U${p}`]
           workbook.sheet(j).cell(hakoo[k]).value(a[k]);
                
                  console.log(a,p)
                  //作りたいbook名を入力
                  workbook.toFileAsync(""); 
                }
           }
      
  
    }
    //cにはシート名を入力
    //buffには読み込むファイルを入力
    var j = 0
    var c = ''
    var buff = fs.readFileSync("");
    var sheet = workbook.sheet(j).name(`${c}`);
    Main()
    
    //cには新たに作りたいシート名を入力
    //buffには読み込むファイルを入力
    var j = 1
    var buff = fs.readFileSync("");
    var c = ''
    var sheet2 = workbook.addSheet('Sheet2', j); // -> Sheet1 Sheet2 Sheet5
    var sheet = workbook.sheet(j).name(`${c}`);
    Main()
    
}); 


  