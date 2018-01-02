(function ($) {

    var defaults = {
        //解析并导出到excel的地址，用来解析传递过去的列集合、当前页数据集合，并最终导出到excel文件
        exportParserHref: "/Test/ExportDataByDataGrid"
    };

    var exportData = function (target, param) {
        var fileType = ["excel"].contains(param.fileType) ? param.fileType : "excel";
        var fileName = param.fileName;
        if (fileType == "excel") { return exportDataToExcel(target, fileName); }
    };

    var exportDataToExcel = function (target, fileName) {
        var t = $(target), dgOptions = t.datagrid("options"), rows = t.datagrid("getRows");
        if (rows.length == 0) { return; }

        var lastFrozenColumns = getColumns(target, true), lastColumns = getColumns(target, false);

        var maxGroup = lastFrozenColumns.length >= lastColumns.length ? lastFrozenColumns.length : lastColumns.length;
        var totalColumns = new Array(maxGroup);
        for (var k = 0; k < maxGroup; k++) {
            if (totalColumns[k] == undefined) { totalColumns[k] = []; }
            if (lastFrozenColumns.length != 0) {
                $.util.merge(totalColumns[k], lastFrozenColumns[k]);
            }
            if (lastColumns.length != 0) {
                $.util.merge(totalColumns[k], lastColumns[k]);
            }
        }
        //用formatter处理rows 为不影响datagrid本身的rows集合，先克隆rows
        var existFields = [], cloneRowsObj = $.extend(true, {}, rows), lastRows = [];
        for (var c in cloneRowsObj) {
            if ($.string.isNumeric(String(c))) { lastRows.push(cloneRowsObj[c]); }
        }

        for (var index = 0; index < totalColumns.length; index++) {
            totalColumns[index].forEach(function (itemCol, itemColIndex) {
                if (itemCol.formatter && $.isFunction(itemCol.formatter)) {
                    lastRows.forEach(function (itemRow, itemRowIndex) {
                        var spe = true, rowIndex = !dgOptions.idField ? itemRowIndex : t.datagrid("getRowIndex", itemRow[dgOptions.idField]);
                        for (var prop in itemRow) {
                            if (prop == itemCol.field) {
                                itemRow[prop] = itemCol.formatter.call(itemCol, itemRow[prop], itemRow, rowIndex);
                                if (!existFields.contains(itemCol.field)) { existFields.push(itemCol.field); }
                                spe = false;
                                break;
                            }
                        }
                        if (spe) {
                            //处理特殊情况，遍历了row的属性，却没找到itemCol.field
                            // field 随便写，通过formatter强行显示指定数据
                            //将field的属性写入row中，其值为formatter的返回值
                            if (itemCol.field && !$.string.isNullOrWhiteSpace(itemCol.field)) {
                                //若该field已存在于 其他列的field 中，则不处理formatter，事实上，easyui也不会对这种formatter进行处理
                                if (!existFields.contains(itemCol.field)) {
                                    itemRow[itemCol.field] = itemCol.formatter.call(itemCol, itemRow[itemCol.field], itemRow, rowIndex);
                                }
                            }
                        }
                    });
                }
            });
        }

        //组装参数
        var param = { data: JSON.stringify(lastRows), columns: JSON.stringify(totalColumns), fileName: fileName == undefined ? "" : fileName };

        //console.log("行数：");
        //console.log(rowspanCount);
        //console.log("********************");
        //console.log("列数：");
        //console.log(colspanCount);
        //console.log("********************");
        //console.log("最终列集合：");
        //console.log(totalColumns);
        //console.log("********************");
        //console.log("最终本页数据集合：");
        //console.log(rows);
        //console.log("********************");
        //console.log("参数：");
        //console.log(param);
        //console.log("********************");

        //模拟form提交完成导出excel操作
        $("iframe[name='hiddenIframe']").remove();
        $("form[target='hiddenIframe']").remove();

        //form的action指向，需结合后台配合，因此本扩展不是真正的easyui扩展。
        var tempForm = $('<form action="' + defaults.exportParserHref + '" target="hiddenIframe" method="post"></form>');
        for (var prop in param) {
            tempForm.append("<input name=\"" + prop + "\" type=\"hidden\" value='" + param[prop] + "' / >");
        }

        $("body").append("<iframe src=\"about:blank\" name=\"hiddenIframe\" style=\"display:none;\"></iframe>").append(tempForm);
        tempForm.submit();
    };

    var getColumns = function (target, frozen) {
        var t = $(target), dgOptions = t.datagrid("options");

        //取目标列集合
        var fColumns = frozen == true ? (dgOptions.frozenColumns || [[]]) : (dgOptions.columns || [[]]).clone();

        //过滤checkbox列和hidden列
        var ddddd = fColumns.clone();
        for (var k = 0; k < fColumns.length; k++) {
            var tempLen = fColumns[k].length;
            for (var inK = 0; inK < tempLen; inK++) {
                var needRemove = false, itemFc = fColumns[k][inK];
                //checkbox列
                if (itemFc.checkbox && itemFc.checkbox == true) { needRemove = true; }
                //无title的列
                if (!needRemove && $.string.isNullOrWhiteSpace(itemFc.title)) { needRemove = true; }
                //hidden列
                if (!needRemove && itemFc.hidden && itemFc.hidden == true) { needRemove = true; }

                if (needRemove) { fColumns[k].removeAt(inK); tempLen--; inK--; }
            }
        }

        //计算列集合的总列数总行数
        //总行数 = fColumns.length
        //总列数 = fColumns[item].colspan之和中最大的
        var colspanCount = $.array.max($.array.map(fColumns, function (itemFc) { return $.array.sum(itemFc, function (item) { return item.colspan || 1; }) })),
            rowspanCount = fColumns.length;

        var lastColumns = [];
        //组建最终列集合的数组结构
        for (var i = 0; i < rowspanCount; i++) {
            lastColumns[i] = new Array(colspanCount);
        }

        var getFixedColumnIndex = function (a) {
            for (var i = 0; i < a.length; i++) {
                if (a[i] == undefined) {
                    return i;
                }
            }
            return -1;
        };
        //往最终列集合里填充数据
        for (var columIndex = 0; columIndex < fColumns.length; columIndex++) {
            fColumns[columIndex].forEach(function (itemFc, itemIndex) {
                var fieldIndex = getFixedColumnIndex(lastColumns[columIndex]); //找到第一个未赋值的元素索引
                if (fieldIndex >= 0) {
                    for (var c = fieldIndex; c < colspanCount ; c++) {
                        var tempCol = $.extend({}, itemFc, {});
                        if (tempCol.colspan == undefined) { tempCol.colspan = 1; }
                        if (tempCol.rowspan == undefined) { tempCol.rowspan = 1; }
                        if ((itemFc.colspan || 1) > 1) {
                            //若列是跨列的，则认为该列的field无效
                            delete tempCol.field;
                        }
                        lastColumns[columIndex][c] = tempCol;
                        if ((itemFc.rowspan || 1) > 1) {
                            for (var d = 1; d < itemFc.rowspan; d++) {
                                if (columIndex + d <= rowspanCount) {
                                    lastColumns[columIndex + d][c] = tempCol;
                                }
                            }
                        }
                        if ((itemFc.colspan || 1) > 1) {
                            for (var d = 1; d < itemFc.colspan; d++) {
                                if (c + d <= colspanCount) {
                                    lastColumns[columIndex][c + d] = tempCol;
                                }
                            }
                        }
                        break;
                    }
                }
            });
        }

        return lastColumns;
    };

    var methods = {

        //  扩展 easyui-datagrid 的自定义方法；导出当前页数据到文件；该方法定义如下参数：
        //      param:  这是一个 JSON-Object 对象，该 JSON-Object 可以包含如下属性：
        //          fileType:        字符串，表示要导出的目标文件类型，其值可以是 excel ，若不传递该参数，则当做 excel ；
        //          fileName:        字符串，表示要导出的目标文件名称，若不传递该参数，则使用默认文件名。
        exportData: function (jq, param) {
            return jq.each(function () {
                exportData(this, param);
            });
        }
    };

    if ($.fn.datagrid.extensions != null && $.fn.datagrid.extensions.methods != null) {
        $.extend($.fn.datagrid.extensions.methods, methods);
    }

    $.extend($.fn.datagrid.methods, methods);

})(jQuery);