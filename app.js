renderGrid()

let isUDSTKDownloaded = false;
let isCDDBDownloaded = false;

function renderGrid(data = []) {
  $('#dx-grid').dxDataGrid({
    dataSource: data,
    // keyExpr: 'Kode Mesin Motor (5 digit pertama)',
    columns,
    showBorders: true,
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

        console.log(STK.length);

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
            'DP': udstk['DP'],
            'Tenor': udstk['Tenor'],
            'Besar Cicilan': udstk['Besar Cicilan'],
            'ID POS': '',
            'No. KTP / NPWP': cddb['No. KTP / NPWP'],
            'Kode Customer': cddb['Kode Customer'],
            'Jenis Kelamin': cddb['Jenis Kelamin'],
            'Tanggal Lahir/ Tanggal pembuatan NPWP': cddb['Tanggal Lahir/ Tanggal pembuatan NPWP'],
            'Alamat Surat': cddb['Alamat Surat'],
            'KODE Kelurahan Surat': cddb['KODE Kelurahan Surat'],
            'KODE Kecamatan Surat': cddb['KODE Kecamatan Surat'],
            'Kota Surat': cddb['Kota Surat'],
            'Kode Pos Surat': cddb['Kode Pos Surat'],
            'Propinsi Surat': cddb['Propinsi Surat'],
            'Agama': cddb['Agama'],
            'Pekerjaan': cddb['Pekerjaan'],
            'Pengeluaran': cddb['Pengeluaran'],
            'Pendidikan': cddb['Pendidikan'],
            'Nama Penanggung Jawab': cddb['Nama Penanggung Jawab'],
            'No.HP (GSM/CDMA)': cddb['No.HP (GSM/CDMA)'],
            'No. Telp': cddb['No. Telp'],
            'Kebersediaan untuk di hub': cddb['Kebersediaan untuk di hub'],
            'Merk Motor yang dimiliki sebelumnya': cddb['Merk Motor yang dimiliki sebelumnya'],
            'Jenis Motor yang dimiliki sebelumnya': cddb['Jenis Motor yang dimiliki sebelumnya'],
            'Sepeda motor yang digunakan untuk': cddb['Sepeda motor yang digunakan untuk'],
            'Yang menggunakan sepeda motor Anda': cddb['Yang menggunakan sepeda motor Anda'],
            'Kode Sales Person': cddb['Kode Sales Person'],
            'Email': cddb['Email'],
            'Status Rumah': cddb['Status Rumah'],
            'Status Nomor HP': cddb['Status Nomor HP'],
            'Facebook': cddb['Facebook'],
            'Twitter': cddb['Twitter'],
            'Instagram': cddb['Instagram'],
            'Youtube (@gmail.com)': cddb['Youtube (@gmail.com)'],
            'Hobi': cddb['Hobi'],
            'Keterangan': cddb['Keterangan '] || '',
            'Kewarganegaraan': cddb['Kewarganegaraan'],
            'No KK': cddb['No KK'],
            'ReferensiID': cddb['ReferensiID'] || '',
            'RO BD ID': cddb['RO BD ID'] || '',
            'Kode FLP Koordinator': cddb['Kode FLP Koordinator'] || '',
            'Seri Mesin RO': cddb['Seri Mesin RO'] || '',
            'No Mesin RO': cddb['No Mesin RO'] || '',
            'TGL SPG': cddb['TGL SPG'] || '',
            'Kota Tempat Tinggal': cddb['Kota Tempat Tinggal'],
            'Nama Instansi': cddb['Nama Instansi'],
            'Kota Instansi': cddb['Kota Instansi'] || '',
            'Kode Kecamatan Instansi': cddb['Kode Kecamatan Instansi'],
            'Kode Kota Instansi': cddb['Kode Kota Instansi'],
            'Kode Propinsi Instansi': cddb['Kode Propinsi Instansi'] || '',
            'Default': '3'
          }

          udstkcdb.push(_obj)
        }

        renderGrid(udstkcdb)

        $('#btnExport').removeClass('d-none')
        $("#btnUDSTK").click(function(e) {
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
                
                downloadBlob(arrayToCsv(_data), `${$('#idkodeahm').val()}-${$('#idkodecabang').val()}-${datestring}11111-FKTR_${$('#idkodempm').val()}001_${_d.getFullYear()}_${_d.getMonth() + 1}_-${_random}.udstk`, 'text/csv;charset=utf-8;')
            })
        })

        $("#btnCDDB").click(function(e) {
          $("#dx-grid").dxDataGrid('instance').getDataSource().store().load().done((res)=>{
            const _data = []
            res.forEach(v => {
              _data.push(
                  [
                  v['Kode Mesin Motor (5 digit pertama)'],
                  v['Nomor Mesin Motor (7 digit terakhir)'],
                  v['No. KTP / NPWP'],
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
                  v['No.HP (GSM/CDMA)'],
                  v['No. Telp'],
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
                  v['Youtube (@gmail.com)'],
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
                
                downloadBlob(arrayToCsv(_data), `${$('#idkodeahm').val()}-${$('#idkodecabang').val()}-${datestring}11111-FKTR_${$('#idkodempm').val()}001_${_d.getFullYear()}_${_d.getMonth() + 1}_-${_random}.cddb`, 'text/csv;charset=utf-8;')
          })
      })

    };
    reader.onerror = function (event) {
        console.error("File could not be read! Code " + event.target.error.code);
    };
    reader.readAsBinaryString(selectedFile);
});

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
  var blob = new Blob([content], { type: contentType + '' });
  var url = URL.createObjectURL(blob);

  // Create a link to download it
  var pom = document.createElement('a');
  pom.href = url;
  pom.setAttribute('download', filename);
  pom.click();
}