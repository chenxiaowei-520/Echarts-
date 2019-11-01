// 数据处理入口
function excelToECharts(obj) {
    excelToData(obj)
}

// 读取Excel转换为json
function excelToData(obj) {
    let files = obj.files;
    // 如果有文件
    if (files.length) {
        // 初始化一个FileReader实例
        let reader = new FileReader();
        let file = files[0];
        // 看下文件是不是xls或者xlsx的
        let fullName = file.name;   // 全名
        let filename = fullName.substring(0, fullName.lastIndexOf("."));    // 文件名
        let fixName = fullName.substring(fullName.lastIndexOf("."), fullName.length);   // 后缀名
        // 处理excel表格
        if (fixName == ".xls" || fixName == ".xlsx") {
            reader.onload = function (ev) {
                let data = ev.target.result;
                // 获取到excel
                let excel = XLSX.read(data, {type: 'binary'});
                // 获取第一个标签页名字
                let sheetName = excel.SheetNames[0];
                // 根据第一个标签页名，获取第一个标签页的内容
                let sheet = excel.Sheets[sheetName];
                // 转换为JSON
                let sheetJson = XLSX.utils.sheet_to_json(sheet);
                
                // 如果有结果，处理结果
                if (sheetJson.length) {
                    // 记录一下各个列名
                    keys = [];
                    for (key in sheetJson[0]) {
                        keys.push(key)
                    }

                    // 处理一下作为x轴的列名和数据
                    let xZhou = {};
                    xZhou.name = keys.splice(0, 1);
                    xDatas = [];
                    for (i in sheetJson) {
                        xDatas.push(sheetJson[i][xZhou.name]);
                    }
                    xZhou.data = xDatas;

                    // 处理主体数据
                    let datas = [];
                    for (i in keys) {
                        let one = {};       // 一组
                        one.name = keys[i]; // 名称
                        one.type = 'line';  // 图表类型
                        one.smooth = true;  // 平滑的线
                        let point = [];     // 记录这一组的所有点
                        for (idx in sheetJson) {
                            // 把这组的点push到数组中
                            point.push(sheetJson[idx][one.name]);
                        }
                        one.data = point;
                        // 把这组数据添加到主体数据中
                        datas.push(one)
                    }

                    // 调用展现的方法
                    dataToEChart(filename, keys, xZhou, datas);

                }
            }
        } else {
            alert("起开，只支持excel")
        }
        reader.readAsBinaryString(file);
    }
}

// 数据展现
function dataToEChart(title, keys, xZhou, datas) {
    // 发现每次执行init的时候会在给的div标签中加入一个_echarts_instance_的属性，后面再init的话就不行了
    // 所以每次先看下有没有这个属性，删掉他
    console.log(document.getElementById('ECharts_main').getAttribute("_echarts_instance_"))
    document.getElementById('ECharts_main').removeAttribute("document.getElementById('_echarts_instance_')");

    // 基于准备好的dom，初始化echarts实例
    var myChart = echarts.init(document.getElementById('ECharts_main'));

    // 指定图表的配置项和数据
    var option = {
        title: {
            text: title,
            x: 'center',
            y: 'bottom'
        },
        tooltip: {
            trigger: 'axis'
        },
        legend: {
            data: keys,
            orient: 'vertical',
            x: 'right',
            y: 'center'
        },
        xAxis: xZhou,
        yAxis: {},
        series: datas,
        toolbox: {
            show: true,
            left: 'right',
            feature: {
                dataView: {},
                magicType: {
                    type: ['line', 'bar', 'stack', 'tiled']
                },
                saveAsImage: {}
            }
        }
    };

    // 使用刚指定的配置项和数据显示图表。
    myChart.setOption(option);
}