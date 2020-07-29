package com.easy.convert.model

/**
 * @author  周泽
 * @date Create in 17:03 2020/7/29
 * @Description
 */
class TableData {

    var headers: List<String>

    private var rows: List<Row>

    var width = 0.0

    var border = 0.0

    fun getRows(): List<Row> {
        return rows
    }

    fun setRows(rows: List<Row>) {
        this.rows = rows
    }

    constructor(headers: List<String>, rows: List<Row>, width: Double, border: Double) : super() {
        this.headers = headers
        this.rows = rows
        this.width = width
        this.border = border
    }

    constructor(headers: List<String>, rows: List<Row>) : super() {
        this.headers = headers
        this.rows = rows
    }

    class Row(var rowData: List<String>)
}