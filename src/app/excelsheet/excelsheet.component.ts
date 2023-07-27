import { Component, OnInit } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-excelsheet',
  templateUrl: './excelsheet.component.html',
  styleUrls: ['./excelsheet.component.css']
})
export class ExcelsheetComponent implements OnInit {

  data: [][];
  cabeceraData: [][];
  indexColCod: number;
  codigosRealizados: any[] = [];
  selectedOption: string;

  valoresJovenMe1er = [
    { codigo: "Z000", labconf: "" },
    { codigo: "C8002", labconf: "1" },
    { codigo: "99403.01", labconf: "1" },
    { codigo: "96150.01", labconf: "" },
    { codigo: "99402.09", labconf: "" },
    { codigo: "Z019", labconf: "DNT" },
    { codigo: "E46X", labconf: "IMC" },
    { codigo: "Z006", labconf: "IMC" },
    { codigo: "E660", labconf: "IMC" },
    { codigo: "E669", labconf: "IMC" },
    { codigo: "U8170", labconf: "RSM" },
    { codigo: "U8170", labconf: "RSA" },
    { codigo: "U8170", labconf: "RMA" },
    { codigo: "99401.16", labconf: "" },
    { codigo: "Z010", labconf: "N" },
    { codigo: "Z011", labconf: "N" },
    { codigo: "99401", labconf: "1" },
    { codigo: "99173", labconf: "20" },
    { codigo: "99173", labconf: "20" },
    { codigo: "Z128", labconf: "N" },
    { codigo: "99402.05", labconf: "" },
    { codigo: "99401.33", labconf: "" },
    { codigo: "86318.01", labconf: "RN" },
    { codigo: "99401.34", labconf: "" },
    { codigo: "87342", labconf: "RN" },
    { codigo: "99199.22", labconf: "N" },
    { codigo: "Z017", labconf: "" },
    { codigo: "85018", labconf: "1" },
    { codigo: "82465", labconf: "" },
    { codigo: "82948", labconf: "" },
    { codigo: "81002", labconf: "" },
    { codigo: "84478", labconf: "" }
  ];
  valoresJovenOtros1er = [
    { codigo: "99402.08", labconf: "1" },
    { codigo: "99208", labconf: "" },
    { codigo: "99402.04", labconf: "1" },
    { codigo: "99402.03", labconf: "1" },
    { codigo: "88141", labconf: "" },
    { codigo: "VACUNA", labconf: "" },
    { codigo: "C0011", labconf: "1" },
    { codigo: "99401", labconf: "" },
    { codigo: "99402.09", labconf: "" }
  ];
  valoresJovenDr2do = [
    { codigo: "Z000", labconf: "" },
    { codigo: "C8002", labconf: "TA" },
    { codigo: "96150.02", labconf: "" },
    { codigo: "99402.09", labconf: "" },
    { codigo: "99403.01", labconf: "2" },
    { codigo: "99401", labconf: "2" },
    { codigo: "U292", labconf: "DNT" }
  ];
  valoresJovenOtros2do = [
    { codigo: "99208", labconf: "" },
    { codigo: "99402.04", labconf: "2" },
    { codigo: "99402.03", labconf: "2" },
    { codigo: "88141", labconf: "N" },
    { codigo: "99402.08", labconf: "2" },
    { codigo: "C0011", labconf: "2" },
    { codigo: "99401", labconf: "" },
    { codigo: "96150.03", labconf: "" },
    { codigo: "99402.09", labconf: "" }
  ];

  valoresAdultoDr1er = [
    { codigo: "Z000", labconf: "" },
    { codigo: "C8002", labconf: "1" },
    { codigo: "Z019", labconf: "DNT" },
    { codigo: "E46X", labconf: "IMC" },
    { codigo: "Z006", labconf: "IMC" },
    { codigo: "E660", labconf: "IMC" },
    { codigo: "E669", labconf: "IMC" },
    { codigo: "U8170", labconf: "RSM" },
    { codigo: "U8170", labconf: "RSA" },
    { codigo: "U8170", labconf: "RMA" },
    { codigo: "99199.22", labconf: "N" },
    { codigo: "99401.13", labconf: "" },
    { codigo: "99401", labconf: "" },
    { codigo: "Z017", labconf: "" },
    { codigo: "82947", labconf: "" },
    { codigo: "84478", labconf: "" },
    { codigo: "82465", labconf: "" },
    { codigo: "85018", labconf: "1" },
    { codigo: "81003", labconf: "" },
    { codigo: "99401", labconf: "1" },
    { codigo: "99403.01", labconf: "1" },
    { codigo: "Z128", labconf: "N" },
    { codigo: "96150.01", labconf: "" },
    { codigo: "99402.09", labconf: "" },
    { codigo: "99401.33", labconf: "" },
    { codigo: "86703.01", labconf: "RN" },
    { codigo: "99401.34", labconf: "" },
    { codigo: "86318.01", labconf: "RN" },
    { codigo: "87342", labconf: "RN" },
    { codigo: "99402.05", labconf: "RN" },
    { codigo: "99173", labconf: "20" },
    { codigo: "99401.16", labconf: "" },
    { codigo: "Z010", labconf: "N" },
    { codigo: "Z011", labconf: "N" },
    { codigo: "84152", labconf: "" },
    { codigo: "99386.03", labconf: "N" },
    { codigo: "82270", labconf: "" }
  ];
  valoresAdultoOtros1er = [
    { codigo: "88141", labconf: "N" },
    { codigo: "99402.08", labconf: "2" },
    { codigo: "96150.03", labconf: "" },
    { codigo: "99402.09", labconf: "2" },
    { codigo: "99402.03", labconf: "2" },
    { codigo: "99402.04", labconf: "2" },
    { codigo: "C0011", labconf: "2" },
    { codigo: "99401", labconf: "" }
  ];
  valoresAdultoDr2do = [
    { codigo: "Z000", labconf: "" },
    { codigo: "C8002", labconf: "TA" },
    { codigo: "99403.01", labconf: "2" },
    { codigo: "U262", labconf: "DNT" },
    { codigo: "84152", labconf: "RN" },
    { codigo: "99401.13", labconf: "" },
    { codigo: "99401", labconf: "2" },
    { codigo: "96150.02", labconf: "2" },
    { codigo: "99402.09", labconf: "2" },
    { codigo: "82270", labconf: "N" },
    { codigo: "U0041", labconf: "" }
  ];
  valoresAdultoOtros2do = [
    { codigo: "88141", labconf: "N" },
    { codigo: "88141.01", labconf: "N" },
    { codigo: "99402.08", labconf: "1" },
    { codigo: "99402.04", labconf: "1" },
    { codigo: "99402.06", labconf: "1" },
    { codigo: "99402.03", labconf: "1" },
    { codigo: "99208", labconf: "" },
    { codigo: "90714", labconf: "1" },
    { codigo: "90658", labconf: "" },
    { codigo: "90749.02", labconf: "" },
    { codigo: "C0011", labconf: "1" },
    { codigo: "99401", labconf: "" },
    { codigo: "96150.04", labconf: "" },
    { codigo: "99402.09", labconf: "" }
  ];

  valoresAdulMayorDr1er = [
    { codigo: "99387", labconf: "AS" },
    { codigo: "Z636.1", labconf: "" },
    { codigo: "C8002", labconf: "1" },
    { codigo: "99401", labconf: "1" },
    { codigo: "E46X", labconf: "IMC" },
    { codigo: "Z006", labconf: "IMC" },
    { codigo: "E660", labconf: "IMC" },
    { codigo: "E669", labconf: "IMC" },
    { codigo: "999403.01", labconf: "1" },
    { codigo: "U8170", labconf: "RSM" },
    { codigo: "U8170", labconf: "RSA" },
    { codigo: "U8170", labconf: "RMA" },
    { codigo: "Z010", labconf: "N" },
    { codigo: "99173", labconf: "20" },
    { codigo: "99401.13", labconf: "" },
    { codigo: "Z011", labconf: "N" },
    { codigo: "99401.13", labconf: "1" },
    { codigo: "96150.01", labconf: "" },
    { codigo: "99402.09", labconf: "" },
    { codigo: "Z017", labconf: "" },
    { codigo: "85018", labconf: "1" },
    { codigo: "82947", labconf: "" },
    { codigo: "82465", labconf: "" },
    { codigo: "84478", labconf: "" },
    { codigo: "81003", labconf: "" },
    { codigo: "99401.33", labconf: "" },
    { codigo: "86318.01", labconf: "RN" },
    { codigo: "99401.34", labconf: "" },
    { codigo: "87342", labconf: "RN" },
    { codigo: "99402.05", labconf: "" },
    { codigo: "Z128", labconf: "N" },
    { codigo: "Z019", labconf: "DNT" },
    { codigo: "99199.22", labconf: "N" },
    { codigo: "99386.03", labconf: "N" },
    { codigo: "Z125", labconf: "" },
    { codigo: "82270", labconf: "" }
  ];
  valoresAdulMayorOtros1er = [
    { codigo: "96150.03", labconf: "" },
    { codigo: "99402.09", labconf: "" },
    { codigo: "88141", labconf: "" },
    { codigo: "99402.08", labconf: "1" },
    { codigo: "90658", labconf: "" },
    { codigo: "C0011", labconf: "1" }
  ];
  valoresAdulMayorDr2do = [
    { codigo: "99387", labconf: "AS" },
    { codigo: "Z636.1", labconf: "" },
    { codigo: "C8002", labconf: "TA" },
    { codigo: "99401", labconf: "2" },
    { codigo: "99403.01", labconf: "2" },
    { codigo: "99401.13", labconf: "2" },
    { codigo: "96150.07", labconf: "" },
    { codigo: "99402.09", labconf: "2" },
    { codigo: "84152", labconf: "N" },
    { codigo: "82270", labconf: "N" }
  ];
  valoresAdulMayorOtros2do = [
    { codigo: "88141", labconf: "" },
    { codigo: "C011", labconf: "2" },
    { codigo: "96150.02", labconf: "" },
    { codigo: "99402.09", labconf: "2" },
    { codigo: "99402.08", labconf: "2" }
  ];

  flagJoven: boolean = false;
  flagAdulto: boolean = false;
  flagAdultoMayor: boolean = false;

  constructor() { }

  ngOnInit(): void {
  }

  onFileChange(evt: any) {
    const target : DataTransfer =  <DataTransfer>(evt.target);

    if (target.files.length !== 1) throw new Error('Cannot use multiple files');

    const reader: FileReader = new FileReader();

    reader.onload = (e: any) => {
      const bstr: string = e.target.result;

      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      const wsname : string = wb.SheetNames[0];

      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      console.log(ws);

      this.data = (XLSX.utils.sheet_to_json(ws, { header: 1 }));
      this.cabeceraData = this.data[1];
      this.data.splice(0, 2);
      console.log("datos:", this.data);
      console.log("datos cabecera:", this.cabeceraData);
      let x = this.data.slice(1);
      console.log(x);

      //Encontrar columna cod
      this.indexColCod = this.cabeceraData.findIndex(e => String(e).toUpperCase()=="CODIGO");
      console.log("indice columna: ",this.indexColCod);

      //Extraer valores de los codigos
      this.data.forEach((e)=>{
        this.codigosRealizados.push({ codigo: this.isUndefined(e[this.indexColCod]), labconf: this.isUndefined(e[this.indexColCod+2]) });
      });
      console.log("codigos encontrados: ",this.codigosRealizados);
        this.flagJoven = true;
        this.flagAdulto = true;
        this.flagAdultoMayor = true;
    };

    reader.readAsBinaryString(target.files[0]);
  }

  cambioSeleccion(){
    console.log("valor select: ",this.selectedOption);
    switch (this.selectedOption) {
      case 'joven':
        this.flagJoven = true;
        this.flagAdulto = false;
        this.flagAdultoMayor = false;
        break;
      case 'adulto':
        this.flagJoven = false;
        this.flagAdulto = true;
        this.flagAdultoMayor = false;
        break;
      case 'adulto-mayor':
        this.flagJoven = false;
        this.flagAdulto = false;
        this.flagAdultoMayor = true;
        break;
      default:
        break;
    }
  }

  comparar(codigo: any): string{
    let var1 = "NO";
    this.codigosRealizados.forEach((e:any) =>{
      if (e.codigo === codigo.codigo) {
        if (e.labconf === codigo.labconf) {
          var1 = 'SI CODIGO y LABCONF'
          if (codigo.labconf === ""){
            var1 = 'SI CODIGO';
          }
        }else{
          var1 = 'SI CODIGO , NO LABCONF ('+e.labconf+')';
        }
      };
    })
    return var1;
  }
  compararJoven2da(codigo: any): string{
    let var1 = "NO";
    let valoresRepe = [
      { codigo: "Z000", contador: 0 },
      { codigo: "C8002", contador: 0 },
      { codigo: "99403.01", contador: 0 },
      { codigo: "99401", contador: 0 },
      { codigo: "99173", contador: 0 },
      { codigo: "99208", contador: 0 },
      { codigo: "99402.04", contador: 0 },
      { codigo: "99402.03", contador: 0 },
      { codigo: "88141", contador: 0 },
      { codigo: "99402.08", contador: 0 },
      { codigo: "C011", contador: 0 },
      { codigo: "99401", contador: 0 },
      { codigo: "99402.09", contador: 0 }
    ];

    this.codigosRealizados.forEach((e:any) =>{
      if (e.codigo === codigo.codigo) {
        if (e.labconf === codigo.labconf) {
          var1 = 'SI CODIGO y LABCONF'
          if (codigo.labconf === ""){
            var1 = 'SI CODIGO';
          }
        }else{
          var1 = 'SI CODIGO , NO LABCONF ('+e.labconf+')';
        }
        var1 += ' , SOLO 1 VEZ';
        valoresRepe.forEach((e) => {
          if(e.codigo == codigo) e.contador += e.contador
          if(e.contador == 2) var1 += ' , SI 2 VECES';
        });
      }
    })
    return var1;
  }

  compararAdulto2da(codigo: any): string{
    let var1 = "NO";
    let valoresRepe = [
      { codigo: "Z000", contador: 0 },
      { codigo: "C8002", contador: 0 },
      { codigo: "99403.01", contador: 0 },
      { codigo: "84152", contador: 0 },
      { codigo: "99401.13", contador: 0 },
      { codigo: "99401", contador: 0 },
      { codigo: "99402.09", contador: 0 },
      { codigo: "82270", contador: 0 },
      { codigo: "88141", contador: 0 },
      { codigo: "99402.08", contador: 0 },
      { codigo: "99402.09", contador: 0 },
      { codigo: "99402.03", contador: 0 },
      { codigo: "99402.04", contador: 0 },
      { codigo: "C0011", contador: 0 },
      { codigo: "99401", contador: 0 }
    ];

    this.codigosRealizados.forEach((e:any) =>{
      if (e.codigo === codigo.codigo) {
        if (e.labconf === codigo.labconf) {
          var1 = 'SI CODIGO y LABCONF'
        }
        if (codigo.labconf !== ""){
          var1 = 'SI CODIGO , NO LABCONF ('+e.labconf+')';
        }
        var1 += ' , SOLO 1 VEZ';
        valoresRepe.forEach((e) => {
          if(e.codigo == codigo) e.contador += e.contador
          if(e.contador == 2) var1 += ' , SI 2 VECES';
        });
      }
    })
    return var1;
  }

  compararAdulMayor2da(codigo: any): string{
    let var1 = "NO";
    let valoresRepe = [
      { codigo: "99387", contador: 0 },
      { codigo: "Z636.1", contador: 0 },
      { codigo: "C8002", contador: 0 },
      { codigo: "99401", contador: 0 },
      { codigo: "99403.01", contador: 0 },
      { codigo: "99401.13", contador: 0 },
      { codigo: "99402.09", contador: 0 },
      { codigo: "82270", contador: 0 },
      { codigo: "88141", contador: 0 },
      { codigo: "C011", contador: 0 },
      { codigo: "99402.09", contador: 0 },
      { codigo: "99402.08", contador: 0 }
    ];

    this.codigosRealizados.forEach((e:any) =>{
      if (e.codigo === codigo.codigo) {
        if (e.labconf === codigo.labconf) {
          var1 = 'SI CODIGO y LABCONF'
        }
        if (codigo.labconf !== ""){
          var1 = 'SI CODIGO , NO LABCONF ('+e.labconf+')';
        }
        var1 += ' , SOLO 1 VEZ';
        valoresRepe.forEach((e) => {
          if(e.codigo == codigo) e.contador += e.contador
          if(e.contador == 2) var1 += ' , SI 2 VECES';
        });
      }
    })
    return var1;
  }

  isUndefined(val: any): string{
    return (val == undefined ? "" : String(val))
  }

}
