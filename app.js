renderGrid()

let isUDSTKDownloaded = false;
let isCDDBDownloaded = false;

function renderGrid(data = []) {
  $('#dx-grid').dxDataGrid({
    dataSource: data,
    // keyExpr: 'Kode Mesin Motor (5 digit pertama)',
    columns,
    showBorders: true,
    allowColumnResizing: true,
    editing: {
        mode: 'row',
        allowUpdating: false,
        allowDeleting: false,
        allowAdding: false,
      },
      scrolling: {
        columnRenderingMode: 'virtual',
      },
      paging: {
        enabled: false,
      },
  });
      
}

  $("#file_upload").change(function (evt) {
    var selectedFile = evt.target.files[0];
    var reader = new FileReader();
    reader.onload = function (event) {
        var data = event.target.result;
        var workbook = XLSX.read(data, {
            type: 'binary'
        });

        const STK = XLSX.utils.sheet_to_row_object_array(workbook.Sheets['STK']);
        const CDB = XLSX.utils.sheet_to_row_object_array(workbook.Sheets['CDB']);
        const udstkcdb = []

        for (let i = 0; i < STK.length; i++) {
          // ['Nomor Rangka Motor', 'Kode Mesin Motor (5 digit pertama)', 'Nomor Mesin Motor (7digit terakhir)', 'Nama Pemilik/ Perusahaan (sesuai KTP/NPWP)', 'Alamat Pemilik/ Perusahaan (sesuai KTP/NPWP)', 'KODE Kelurahan Pemilik/ Perusahaan (sesuai KTP/NPWP)', 'KODE Kecamatan Pemilik/ Perusahaan (sesuai KTP/ NPWP)', 'Kota Pemilik/ Perusahaan (sesuai KTP/ NPWP)', 'Kode Pos(sesuai KTP/NPWP)', 'Propinsi (sesuai KTP/NPWP)', 'Jenis Pembayaran', 'Kode Dealer ', 'Kode Fincoy', 'DP', 'Tenor', 'Besar Cicilan']
          
          // ['Kode Mesin Motor (5 digit pertama)', 'Nomor Mesin Motor (7digit terakhir)', 'No. KTP / NPWP', 'Kode Customer', 'Jenis Kelamin', 'Tanggal Lahir/ Tanggal pembuatan NPWP', 'Alamat Surat', 'KODE Kelurahan Surat', 'KODE Kecamatan Surat', 'Kota Surat', 'Kode Pos Surat', 'Propinsi Surat', 'Agama', 'Pekerjaan', 'Pengeluaran', 'Pendidikan', 'Nama Penanggung Jawab', 'No.HP (GSM/CDMA)', 'No. Telp', 'Kebersediaan untuk di hub', 'Merk Motor yang dimiliki sebelumnya', 'Jenis Motor yang dimiliki sebelumnya', 'Sepeda motor yang digunakan untuk', 'Yang menggunakan sepeda motor Anda', 'Kode Sales Person', 'Email', 'Status Rumah', 'Status Nomor HP', 'Facebook', 'Twitter', 'Instagram', 'Youtube (@gmail.com)', 'Hobi', 'Kewarganegaraan', 'No KK', 'Kota Tempat Tinggal', 'Nama Instansi', 'Kota Instansi', 'Kode Kecamatan Instansi', 'Kode Kota Instansi', 'Kode Propinsi Instansi', 'Default : 3']
          // console.log(Object.keys(STK[i]));
          // console.log(Object.keys(CDB[i]));

          const udstk = STK[i];
          const cddb = CDB[i];
          

          const _obj = {
            'Nomor Rangka Motor': udstk['Nomor Rangka Motor'],
            'Kode Mesin Motor (5 digit pertama)': udstk['Kode Mesin Motor (5 digit pertama)'],
            'Nomor Mesin Motor (7 digit terakhir)': udstk['Nomor Mesin Motor (7digit terakhir)'],
            'Nama Pemilik/ Perusahaan (sesuai KTP/NPWP)': udstk['Nama Pemilik/ Perusahaan (sesuai KTP/NPWP)'],
            'Alamat Pemilik/ Perusahaan (sesuai KTP/NPWP)': udstk['Alamat Pemilik/ Perusahaan (sesuai KTP/NPWP)'],
            'KODE Kelurahan Pemilik/ Perusahaan (sesuai KTP/NPWP)': udstk['KODE Kelurahan Pemilik/ Perusahaan (sesuai KTP/NPWP)'],
            'KODE Kecamatan Pemilik/ Perusahaan (sesuai KTP/ NPWP)': udstk['KODE Kecamatan Pemilik/ Perusahaan (sesuai KTP/ NPWP)'],
            'Kota Pemilik/ Perusahaan (sesuai KTP/ NPWP)': udstk['Kota Pemilik/ Perusahaan (sesuai KTP/ NPWP)'],
            'Kode Pos(sesuai KTP/NPWP)': udstk['Kode Pos(sesuai KTP/NPWP)'],
            'Propinsi (sesuai KTP/NPWP)': udstk['Propinsi (sesuai KTP/NPWP)'],
            'Jenis Pembayaran': udstk['Jenis Pembayaran'],
            'Kode Dealer': udstk['Kode Dealer '],
            'Kode Fincoy': udstk['Kode Fincoy'],
            'DP': udstk['DP'] || 0,
            'Tenor': udstk['Tenor'] || 0,
            'Besar Cicilan': udstk['Besar Cicilan'] || 0,
            'ID POS': '',
            'NoKTP': cddb['No. KTP / NPWP'],
            'Kode Customer': cddb['Kode Customer'],
            'Jenis Kelamin': cddb['Jenis Kelamin'],
            'Tanggal Lahir/ Tanggal pembuatan NPWP': cddb['Tanggal Lahir/ Tanggal pembuatan NPWP'],
            'Alamat Surat': cddb['Alamat Surat'],
            'KODE Kelurahan Surat': cddb['KODE Kelurahan Surat'],
            'KODE Kecamatan Surat': cddb['KODE Kecamatan Surat'],
            'Kota Surat': cddb['Kota Surat'],
            'Kode Pos Surat': cddb['Kode Pos Surat'] || '60000',
            'Propinsi Surat': cddb['Propinsi Surat'],
            'Agama': cddb['Agama'],
            'Pekerjaan': cddb['Pekerjaan'],
            'Pengeluaran': cddb['Pengeluaran'],
            'Pendidikan': cddb['Pendidikan'],
            'Nama Penanggung Jawab': cddb['Nama Penanggung Jawab'] || 'N',
            'NoHP': cddb['No.HP (GSM/CDMA)'] || '0',
            'NoTelp': cddb['No. Telp'] || '0',
            'Kebersediaan untuk di hub': cddb['Kebersediaan untuk di hub'] || 'Y',
            'Merk Motor yang dimiliki sebelumnya': cddb['Merk Motor yang dimiliki sebelumnya'],
            'Jenis Motor yang dimiliki sebelumnya': cddb['Jenis Motor yang dimiliki sebelumnya'],
            'Sepeda motor yang digunakan untuk': cddb['Sepeda motor yang digunakan untuk'],
            'Yang menggunakan sepeda motor Anda': cddb['Yang menggunakan sepeda motor Anda'],
            'Kode Sales Person': cddb['Kode Sales Person'],
            'Email': cddb['Email'],
            'Status Rumah': cddb['Status Rumah'],
            'Status Nomor HP': cddb['Status Nomor HP'],
            'Facebook': cddb['Facebook'] || 'N',
            'Twitter': cddb['Twitter'] || 'N',
            'Instagram': cddb['Instagram'] || 'N',
            'Youtube': cddb['Youtube (@gmail.com)'] || 'N',
            'Hobi': cddb['Hobi'],
            'Keterangan': cddb['Keterangan '] || '',
            'Kewarganegaraan': (!cddb['Kewarganegaraan'] || isNaN(cddb['Kewarganegaraan'])) ? '1' : cddb['Kewarganegaraan'],
            'No KK': cddb['No KK'],
            'ReferensiID': cddb['ReferensiID'] || '',
            'RO BD ID': cddb['RO BD ID'] || '',
            'Kode FLP Koordinator': cddb['Kode FLP Koordinator'] || '',
            'Seri Mesin RO': cddb['Seri Mesin RO'] || '',
            'No Mesin RO': cddb['No Mesin RO'] || '',
            'TGL SPG': cddb['TGL SPG'] || '',
            'Kota Tempat Tinggal': cddb['Kota Tempat Tinggal'] || '',
            'Nama Instansi': cddb['Nama Instansi'] || '',
            'Kota Instansi': cddb['Kota Instansi'] || '',
            'Kode Kecamatan Instansi': cddb['Kode Kecamatan Instansi'] || '',
            'Kode Kota Instansi': cddb['Kode Kota Instansi'] || '',
            'Kode Propinsi Instansi': cddb['Kode Propinsi Instansi'] || '',
            'Default': '3'
          }
          console.log(_obj);
          udstkcdb.push(_obj)
        }

        const _errData = validateData(udstkcdb)

        if (_errData
        && Object.keys(_errData).length === 0
        && Object.getPrototypeOf(_errData) === Object.prototype) {  
          showExportBtn(udstkcdb)
        } else {
          showErrBtn(_errData)
        }

        

    };
    reader.onerror = function (event) {
        console.error("File could not be read! Code " + event.target.error.code);
    };
    reader.readAsBinaryString(selectedFile);
});

function showExportBtn(data) {
  renderGrid(data)
  $('#btnErrContainer').addClass('d-none')
  $('#btnExport').removeClass('d-none')
}

function showErrBtn(data) {
  $('#btnExport').addClass('d-none')
  $('#btnErrContainer').removeClass('d-none')

  $("#btnErr").click(function(e) {
    var wb = XLSX.utils.book_new();
    wb.Props = {
        Title: "Keterangan Error CDDB UDSTK",
        CreatedDate: new Date()
    };
    wb.SheetNames.push("Keterangan Error");
    const ws_data = [...Object.values(data)];
    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    wb.Sheets["Keterangan Error"] = ws;

    const wbout = XLSX.write(wb, {bookType:'xlsx',  type: 'binary'});
    
    downloadBlob(s2ab(wbout), 'keterangan error.xlsx', 'application/octet-stream')
  })
}

function validateData(data) {

  const errors = {}
  
  data.forEach((val, idx) => {
    let errData = []

    if (isNaN(val['KODE Kelurahan Pemilik/ Perusahaan (sesuai KTP/NPWP)'])) {
      errData.push("kelurahan harus berupa kode (Angka)")
    }

    if (isNaN(val['KODE Kecamatan Pemilik/ Perusahaan (sesuai KTP/ NPWP)'])) {
      errData.push("kecamatan harus berupa kode (Angka)")
    }

    if (isNaN(val['Kota Pemilik/ Perusahaan (sesuai KTP/ NPWP)'])) {
      errData.push("Kota harus berupa kode (Angka)")
    }

    if (isNaN(val['Propinsi (sesuai KTP/NPWP)'])) {
      errData.push("Propinsi harus berupa kode (Angka)")
    }

    if (isNaN(val['Jenis Pembayaran'])) {
      errData.push("Jenis Pembayaran harus berupa kode: 1 (Cash), 2 (Credit)")
    }

    if (isNaN(val['Jenis Kelamin']) && val['Jenis Kelamin'] != 'N') {
      errData.push("Jenis Kelamin harus berupa kode: 1 - Laki-laki, 2 - Perempuan, N - Untuk Group")
    }

    if (isNaN(val['Tanggal Lahir/ Tanggal pembuatan NPWP']) && val['Tanggal Lahir/ Tanggal pembuatan NPWP'].length != 8) {
      errData.push("Tanggal Lahir harus berformat DDMMYYY, misal tanggal 05-10-2000 berarti harus diisi 05102000")
    }

    if(isNaN(val['Merk Motor yang dimiliki sebelumnya'])) {
      errData.push('Merk Motor yang dimiliki sebelumnya harus berupa Kode (angka)')
    }

    if(isNaN(val['Jenis Motor yang dimiliki sebelumnya'])) {
      errData.push('Jenis Motor yang dimiliki sebelumnya harus berupa Kode (angka)')
    }

    if(isNaN(val['Sepeda motor yang digunakan untuk'])) {
      errData.push('Sepeda motor yang digunakan untuk harus berupa Kode (angka)')
    }

    if(isNaN(val['Yang menggunakan sepeda motor Anda'])) {
      errData.push('Yang menggunakan sepeda motor Anda harus berupa Kode (angka)')
    }

    if(isNaN(val['Kode Sales Person'])) {
      errData.push('Kode Sales Person harus berupa Kode (angka)')
    }

    if(isNaN(val['Status Rumah']) && val['Status Rumah'] != 'N') {
      errData.push('Status Rumah harus berupa kode sesuai di masterdata excel')
    }

    if(isNaN(val['Status Nomor HP']) && val['Status Nomor HP'] != 'N') {
      errData.push('Status Nomor HP harus berupa kode sesuai di masterdata excel')
    }

    if(isNaN(val['Kode Dealer'])) {
      errData.push('Kode Dealer harus berupa Kode Dealer AHM (6 digit angka)')
    }

    if(!masterdata['hobi'].includes(val['Hobi'])) {
      errData.push('Hobi harus berupa kode sesuai di masterdata excel')
    }

    if(!masterdata['fincoy'].includes(val['Kode Fincoy'])) {
      errData.push('Kode Fincoy harus berupa kode sesuai di masterdata excel')
    }


    if(isNaN(val['Kode Kecamatan Instansi'])) {
      errData.push('Kode Kecamatan Instansi harus berupa Kode (angka)')
    }
    if(isNaN(val['Kode Kota Instansi'])) {
      errData.push('Kode Kota Instansi harus berupa Kode (angka)')
    }
    if(isNaN(val['Kode Propinsi Instansi'])) {
      errData.push('Kode Propinsi Instansi harus berupa Kode (angka)')
    }

    if (errData.length != 0) {
      errData = [`Baris ke-${(idx + 1)}`, ...errData]
      errors[idx] = errData
    }

  })

  return errors

}

function arrayToCsv(data){
  return data.map(row =>
    row
    .map(String)  // convert every value to String
    // .map(v => v.replaceAll('"', '""'))  // escape double colons
    // .map(v => `"${v}"`)  // quote it
    .join(';')  // comma-separated
  ).join('\r\n');  // rows starting on new lines
}

function downloadBlob(content, filename, contentType) {
  // Create a blob
  var blob = new Blob([content], { type: contentType });
  var url = URL.createObjectURL(blob);

  // Create a link to download it
  var pom = document.createElement('a');
  pom.href = url;
  pom.setAttribute('download', filename);
  pom.click();
}


$("#btnUDSTK").click(function(e) {
  const err = validateInput()

  if(err.length != 0) {
    alert(err.join('\n'))
    return
  }

  $("#dx-grid").dxDataGrid('instance').getDataSource().store().load().done((res)=>{
      const _data = []
      res.forEach(v => {
        _data.push(
            [
            v['Nomor Rangka Motor'],
            v['Kode Mesin Motor (5 digit pertama)'],
            v['Nomor Mesin Motor (7 digit terakhir)'],
            v['Nama Pemilik/ Perusahaan (sesuai KTP/NPWP)'],
            v['Alamat Pemilik/ Perusahaan (sesuai KTP/NPWP)'],
            v['KODE Kelurahan Pemilik/ Perusahaan (sesuai KTP/NPWP)'],
            v['KODE Kecamatan Pemilik/ Perusahaan (sesuai KTP/ NPWP)'],
            v['Kota Pemilik/ Perusahaan (sesuai KTP/ NPWP)'],
            v['Kode Pos(sesuai KTP/NPWP)'],
            v['Propinsi (sesuai KTP/NPWP)'],
            v['Jenis Pembayaran'],
            v['Kode Dealer'],
            v['Kode Fincoy'],
            v['DP'],
            v['Tenor'],
            v['Besar Cicilan'],
            v['ID POS'],
            ''
          ]
        )
      })

      let rand = Date.now().toString();
      rand = rand.substring(rand.length - 5)

      let _random = rand
      if (!sessionStorage.getItem('rand')) {
        sessionStorage.setItem('rand', rand)
      } else {
        _random = sessionStorage.getItem('rand')
      }

      const _d = new Date()
      const datestring = _d.getFullYear() + "-" + ("0"+(_d.getMonth()+1)).slice(-2) +"-"+("0" + _d.getDate()).slice(-2);

      isUDSTKDownloaded = true
      if (isCDDBDownloaded) {
        isUDSTKDownloaded = false
        isCDDBDownloaded = false
        sessionStorage.removeItem('rand')
      }
      
      downloadBlob(arrayToCsv(_data), `${$('#idkodeahm').val().toUpperCase()}-${$('#idkodecabang').val().toUpperCase()}-${datestring}11111-FKTR_${$('#idkodempm').val().toUpperCase()}001_${_d.getFullYear()}_${_d.getMonth() + 1}_-${_random}.udstk`, 'text/csv;charset=utf-8;')
  })
})

function s2ab(s) { 
  var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
  var view = new Uint8Array(buf);  //create uint8array as viewer
  for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; //convert to octet
  return buf;    
}

$("#btnCDDB").click(function(e) {
  const err = validateInput()

  if(err.length != 0) {
    alert(err.join('\n'))
    return
  }

  $("#dx-grid").dxDataGrid('instance').getDataSource().store().load().done((res)=>{
    const _data = []
    res.forEach(v => {
      _data.push(
          [
          v['Kode Mesin Motor (5 digit pertama)'],
          v['Nomor Mesin Motor (7 digit terakhir)'],
          v['NoKTP'],
          v['Kode Customer'],
          v['Jenis Kelamin'],
          v['Tanggal Lahir/ Tanggal pembuatan NPWP'],
          v['Alamat Surat'],
          v['KODE Kelurahan Surat'],
          v['KODE Kecamatan Surat'],
          v['Kota Surat'],
          v['Kode Pos Surat'],
          v['Propinsi Surat'],
          v['Agama'],
          v['Pekerjaan'],
          v['Pengeluaran'],
          v['Pendidikan'],
          v['Nama Penanggung Jawab'],
          v['NoHP'],
          v['NoTelp'],
          v['Kebersediaan untuk di hub'],
          v['Merk Motor yang dimiliki sebelumnya'],
          v['Jenis Motor yang dimiliki sebelumnya'],
          v['Sepeda motor yang digunakan untuk'],
          v['Yang menggunakan sepeda motor Anda'],
          v['Kode Sales Person'],
          v['Email'],
          v['Status Rumah'],
          v['Status Nomor HP'],
          v['Facebook'],
          v['Twitter'],
          v['Instagram'],
          v['Youtube'],
          v['Hobi'],
          v['Keterangan'],
          v['Kewarganegaraan'],
          v['No KK'],
          v['ReferensiID'],
          v['RO BD ID'],
          v['Kode FLP Koordinator'],
          v['Seri Mesin RO'],
          v['No Mesin RO'],
          v['TGL SPG'],
          v['Kota Tempat Tinggal'],
          v['Nama Instansi'],
          v['Kota Instansi'],
          v['Kode Kecamatan Instansi'],
          v['Kode Kota Instansi'],
          v['Kode Propinsi Instansi'],
          v['Default'],
          '',
          ''
        ]
      )
    })

    let rand = Date.now().toString();
        rand = rand.substring(rand.length - 5)

        let _random = rand
        if (!sessionStorage.getItem('rand')) {
          sessionStorage.setItem('rand', rand)
        } else {
          _random = sessionStorage.getItem('rand')
        }

        const _d = new Date()
        const datestring = _d.getFullYear() + "-" + ("0"+(_d.getMonth()+1)).slice(-2) +"-"+("0" + _d.getDate()).slice(-2);

        isCDDBDownloaded = true
        if (isUDSTKDownloaded) {
          isUDSTKDownloaded = false
          isCDDBDownloaded = false
          sessionStorage.removeItem('rand')
        }
        
        downloadBlob(arrayToCsv(_data), `${$('#idkodeahm').val().toUpperCase()}-${$('#idkodecabang').val().toUpperCase()}-${datestring}11111-FKTR_${$('#idkodempm').val().toUpperCase()}001_${_d.getFullYear()}_${_d.getMonth() + 1}_-${_random}.cddb`, 'text/csv;charset=utf-8;')
  })
})

const masterdata = {
  'hobi': ['A1','A10','A11','A12','A13','A14','A15','A16','A17','A18','A19','A2','A20','A21','A22','A23','A24','A25','A26','A27','A28','A29','A3','A30','A31','A32','A33','A4','A5','A6','A7','A8','A9'],
  'fincoy': ['ADIRA','BCAMF','FIFIN','MCFIN','MEGAF','MPMFI','MUFIN','NUSAC','OTHER','SUMIT','TUNAI','WOMFN']
}

function validateInput() {
  const err = []
  
  const idkodeahm = $('#idkodeahm').val()
  if (idkodeahm == '') {
    err.push('Kode AHM harus diisi')
  } else if (isNaN(idkodeahm) || idkodeahm.length != 5) {
    err.push('Kode AHM harus berupa angka 5 digit, contoh: 00006')
  }

  const idkodecabang = $('#idkodecabang').val() 
  if (idkodecabang == '') {
    err.push('Kode Cabang harus diisi')
  } else if (idkodecabang.toUpperCase() != 'M2Z' && idkodecabang.toUpperCase() != 'M3Z') {
    err.push('Kode Cabang harus berisi M2Z atau M3Z')
  }

  const idkodempm = $('#idkodempm').val() 
  if (idkodempm == '') {
    err.push('Kode MPM harus diisi')
  } else if (!isNaN(idkodempm[0]) || idkodempm.length != 5) {
    err.push('Kode MPM harus berupa angka 5 digit, contoh: A0000')
  }
  
  return err
  
}