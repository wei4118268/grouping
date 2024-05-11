$(document).ready(() => {

    excelArr = [];

    $('#upload').click(() => {
        $('#upload_input').click();
    })
    $('#upload_input').on('change', e => {
        excelArr = [];
        let file = e.target.files[0];
        if(file){
            //读取excel
            analysis(file).then(tableJson => {
                console.log(tableJson);
                for(v of tableJson){
                    excelArr.push(v.sheet);
                }
                //获取核心数据
                console.log(excelArr);
                //显示开始分组按钮
                $('.grouping').show();
                $('.t2').text('文件读取完成，请点击进行分组')
            })
        }
    })

    $('.btn').click(() => {
        var r = [['选手1', '选手2']];
        var temp = [];

        $('.t2').text('分组中，请稍后...');

        result = group(excelArr);

        console.log(result);
        result = shuffle(result);
        console.log(result);
        

        $('.t2').text('分组完成~等待下载excel文件')

        //result格式化
        for(let i = 0; i < result.length; i++){
            temp = [];
            temp.push(JSON.stringify(result[i][0]).replaceAll('"', '').slice(1, -1));
            if(result[i].length > 1){
                temp.push(JSON.stringify(result[i][1]).replaceAll('"', '').slice(1, -1));
            }
            
            r.push(temp);
            r.push([null, null])
        }

        console.log(r)

        var ws = XLSX.utils.aoa_to_sheet(r);
        ws['!cols'] = [{wpx: 200}, {wpx: 200}]
        console.log(ws)
        var blob = sheet2blob(ws, '分组结果');
        openDownloadXLSXDialog(blob, '123.xlsx');
    })
})

function analysis (file){
    return new Promise(function (resolve) {
        const reader =new FileReader()
        reader.onload = function (e){
            const data = e.target.result
            const datajson = XLSX.read(data, {type: "binary"})
            const result = []
            datajson.SheetNames.forEach(sheetName => {
                result.push({
                    sheetName: sheetName,
                    sheet: XLSX.utils.sheet_to_json(datajson.Sheets[sheetName])
                })
            })
            resolve(result)
        }
        reader.readAsBinaryString(file);
    });
}

function group(arr_arr){
    var temp_arr = [...arr_arr];
    
    var longest = [], longest_index = 0;
    var others = [];
    var others2 = [];

    var result = [];

    var t1, t2, tx = 0;


    if(temp_arr.length == 1){
        temp_arr = [...temp_arr[0]];
        temp_arr = shuffle(temp_arr);

        for(let i = 0; i < temp_arr.length; i = i + 2){
            if(i + 1 == temp_arr.length){
                //out of range
                result.push([temp_arr[i]]);
            }else{
                result.push([temp_arr[i], temp_arr[i + 1]]);
            }
        }

        return result;
    }


    //find longest
    for(let i = 0; i < temp_arr.length; i++){
        if(temp_arr[i].length > longest.length){
            longest = [...temp_arr[i]];
            longest_index = i;
        }
    }

    temp_arr.splice(longest_index, 1);
    others = [...temp_arr];

    //start grouping
    //shuffle
    longest = shuffle(longest);
    others = shuffle(others);
    for(let i = 0; i < others.length; i++){
        others[i] = shuffle(others[i]);
    }
    
    while(longest.length > 0){
        t1 = null; t2 = null;

        if(longest.length == 0){
            //longest is over
            break;
        }

        t1 = longest.splice(0, 1)[0];
        
        //get one from line in others
        for(let i = tx; i < others.length; i++){
            if(others[i].length > 0){
                t2 = others[i].splice(0, 1)[0];
                tx = i + 1;
                if(tx == others.length){
                    tx = 0;
                    //shuffle
                    others = shuffle(others);
                }
                break;
            }
        }

        if(t2 == null){
            //others is over
            longest.push(t1);
            break;
        }

        result.push(shuffle([t1, t2]));
    }

    if(longest.length > 0){
        result = [...result, ...group([longest])]
    }else{
        //delete empty array in others
        for(let i = 0; i < others.length; i++){
            if(others[i].length != 0){
                others2.push(others[i]);
            }
        }
        if(others2.length > 0){
            result = [...result, ...group(others2)]
        } 
    }
    
    return result;
}

function shuffle(arr){
    var result = [],
        random;
    while(arr.length>0){
        random = Math.floor(Math.random() * arr.length);
        result.push(arr[random])
        arr.splice(random, 1)
    }
    return result;
}

function sheet2blob(sheet, sheetName){
    sheetName = sheetName || 'sheet1';
    var workbook = {
        SheetNames: [sheetName],
        Sheets: {}
    };
    workbook.Sheets[sheetName] = sheet;
    // 生成excel的配置项
    var wopts = {
        bookType: 'xlsx', // 要生成的文件类型
        bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        type: 'binary'
    };
    var wbout = XLSX.write(workbook, wopts);
    var blob = new Blob([s2ab(wbout)], {type:"application/octet-stream"});
    // 字符串转ArrayBuffer
    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
    return blob;
}

function openDownloadXLSXDialog(url, saveName){
    if(typeof url == 'object' && url instanceof Blob){
        url = URL.createObjectURL(url); // 创建blob地址
    }
    var aLink = document.createElement('a');
    aLink.href = url;
    aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
    var event;
    if(window.MouseEvent) event = new MouseEvent('click');
    else{
        event = document.createEvent('MouseEvents');
        event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
    }
    aLink.dispatchEvent(event);
}
