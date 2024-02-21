## PHPExport
Library tambahan untuk Codeigniter V3, yang berfungsi untuk membuat export excel tanpa ribet.

### Requirement
-  PHPSpreadsheet
-  Data array yang akan dieksport


### Usage


#### Tambahkan file ke library CI lalu Load librarynya
```
$this->load->library('PHPExport');
```

#### Panggil classnya
```
$exportExcel= new PHPExport; 			
$exportExcel
  ->dataSet($data_set)                         : mandatory
  ->rataTengah('4,5')                          : optional (untuk rata tengah field, isikan nomor kolom)
  ->rataKanan('13')                            : optional (untuk rata kanan field, isikan nomor kolom)
  ->warnaHeader('555555','FFFFFF')             : optional (untuk warna header dan warna font, RBG value)
  ->excel2003('Laporan-SPK_'.date('YmdHis'));  : mandatory (excel2003/excel2007/csv, isikan nama filenya)
```

### Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request
