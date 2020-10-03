import React, { Component } from "react";
import * as Excel from "exceljs";
import { saveAs } from 'file-saver';
import { ExcelRenderer } from "react-excel-renderer";

class HandleExcel extends Component {
  constructor(props) {
    super(props)
    this.state = {
      step: 0,
      table: []
    }
  }

  componentDidMount() {
  }

  componentDidUpdate(prevProps) {
  }

  _ImportExcelSended = async (event) => {
    let fileObj = event.target.files[0];
    ExcelRenderer(fileObj, (err, resp) => {
      if (err) {
        console.log(err);
      } else {
        this.setState({
          colsSended: resp.cols,
          rowsSended: resp.rows
        })
      }
    });
    event.target.value = null;
    this.state.step = 1
    this.forceUpdate()
  }

  _ImportExcelReceivedMoney = async (event) => {
    let fileObj = event.target.files[0];
    //just pass the fileObj as parameter
    ExcelRenderer(fileObj, (err, resp) => {
      if (err) {
        console.log(err);
      } else {
        this.setState({
          cowsReceivedMoney: resp.cols,
          rowsReceivedMoney: resp.rows
        })
      }
    });
    event.target.value = null;
    this.state.step = 2
    this.forceUpdate()
  }

  _handleExportExcel = async (table) => {
    const wb = new Excel.Workbook();
    const ws = wb.addWorksheet('DsBanGhi');

    ws.addRows(table);

    ws.getRow(1).font = { name: 'Times New Roman', family: 2, size: 10, bold: true };

    for (let i = 0; i < table.length + 1; i++) {
      for (let j = 0; j < 7; j++) {
        ws.getCell(String.fromCharCode(65 + j) + (i + 1)).border = {
          top: { style: 'thin', color: { argb: '00000000' } },
          left: { style: 'thin', color: { argb: '00000000' } },
          bottom: { style: 'thin', color: { argb: '00000000' } },
          right: { style: 'thin', color: { argb: '00000000' } }
        }
      }
    }

    wb.xlsx.writeBuffer().then((data) => {
      const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8' });
      saveAs(blob, 'ĐƠN CHƯA CÓ TIỀN.xlsx');
    });
  }

  dataTable = () => {
    let { rowsSended, rowsReceivedMoney } = this.state
    let table = []
    let row = ["TÊN KHÁCH", "SỐ ĐT", "ĐỊA CHỈ", "SẢN PHẨM", "TIỀN", "LƯU Ý", "TỔNG TIỀN ĐI"]
    table.push(row)
    rowsSended.forEach((item, index) => {
      if (item.length > 0 && index > 0) {
        let flag = false
        rowsReceivedMoney.forEach((itemRM, indexRM) => {
          let nameSended = item[0]
          let nameRM = itemRM[4]
          if (nameSended === nameRM) {
            flag = true
          }
        })
        if (flag === false) {
          table.push(item)
        }
      }
    })
    return table
  }

  _handleCheck = async () => {
    let table = this.dataTable()
    this.state.table = table
    this.forceUpdate()
  }

  _handleDownload = async () => {
    let table = this.dataTable()
    setTimeout(() => {
      this._handleExportExcel(table)
    }, 500)
  }

  render() {
    let { step, table } = this.state
    return (
      <div className='container'>
        <div className="card-header">
          <button className='btn btn-sm'>
            <span>
              <label htmlFor="file-upload1" className="btn btn-sm btn-outline-primary border-radius">
                <i className="fas fa-upload"></i>Import đơn đã gửi
                </label>
              <input
                id="file-upload1"
                className="btn btn-sm btn-outline-primary border-radius"
                type="file"
                value={this.state.value}
                onChange={this._ImportExcelSended.bind(this)}
                style={{ display: 'none' }}
              />
            </span>
          </button>

          <button className='btn btn-sm'>
            <span>
              <label htmlFor="file-upload2" className="btn btn-sm btn-outline-primary border-radius">
                <i className="fas fa-upload"></i>Import đơn đã có tiền
                </label>
              <input
                id="file-upload2"
                className="btn btn-sm btn-outline-primary border-radius"
                type="file"
                value={this.state.value}
                onChange={this._ImportExcelReceivedMoney.bind(this)}
                style={{ display: 'none' }}
              />
            </span>
          </button>

          <button className='btn btn-sm' onClick={() => this._handleCheck()}>
            <label className="btn btn-sm btn-outline-primary border-radius">
              Kiểm tra đơn chưa có tiền
          </label>
          </button>

          <button className='btn btn-sm' onClick={() => this._handleDownload()}>
            <label className="btn btn-sm btn-outline-primary border-radius">
              Tải file Excel
          </label>
          </button>
        </div>
        <div className="card-body">
          {step >= 0 && <React.Fragment>1. Import đơn đã gửi<br /></React.Fragment>}
          {step >= 1 && <React.Fragment>2. Import đơn đã có tiền<br /></React.Fragment>}
          {step >= 2 && <React.Fragment>3. Kiểm tra danh sách đơn hoặc Xuất file excel đơn chưa có tiền</React.Fragment>}
        </div>
        <div className="card-body">
          <table className="table table-bordered">
            <tbody>
              {table.map((item1, ind1) => {
                return (
                  <tr key={ind1}>
                    <td>{ind1}</td>
                    {item1.map((item2, ind2) => {
                      return <td key={ind2}>{item2}</td>
                    })}
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>

      </div>
    );
  }

}

export default HandleExcel;
