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
    { codigo: "Z000", color:"#D2BF55",  labconf: "" },
    { codigo: "C8002", color:"#FFEED6", labconf: "1" },
    { codigo: "99403.01", color:"#4CE0B3", labconf: "1" },
    { codigo: "96150.01", labconf: "" },
    { codigo: "99402.09", color:"#FF8E72", labconf: "" },
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
    { codigo: "99401", color:"#85CB33", labconf: "1" },
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
    { codigo: "99402.08", color:"#E16036", labconf: "1" },
    { codigo: "99208", color:"#8CA0D7", labconf: "" },
    { codigo: "99402.04", color:"#9D79BC", labconf: "1" },
    { codigo: "99402.03", color:"#91C4F2", labconf: "1" },
    { codigo: "88141", color:"#D6A99A", labconf: "" },
    { codigo: "VACUNA", labconf: "" },
    { codigo: "C0011", color:"#C8AB83", labconf: "1" },
    { codigo: "99401", color:"#85CB33", labconf: "" },
    { codigo: "96150.04", labconf: "" },
    { codigo: "99402.09", color:"#FF8E72", labconf: "" }
  ];
  valoresJovenDr2do = [
    { codigo: "Z000", color:"#D2BF55",  labconf: "" },
    { codigo: "C8002", color:"#FFEED6", labconf: "TA" },
    { codigo: "96150.02", color:"#FBBFCA", labconf: "" },
    { codigo: "99402.09", color:"#FF8E72", labconf: "" },
    { codigo: "99403.01", color:"#4CE0B3", labconf: "2" },
    { codigo: "99401", color:"#85CB33", labconf: "2" },
    { codigo: "U292", color:"#F7D002", labconf: "DNT" }
  ];
  valoresJovenOtros2do = [
    { codigo: "99208", color:"#8CA0D7", labconf: "" },
    { codigo: "99402.04", color:"#9D79BC", labconf: "2" },
    { codigo: "99402.03", color:"#91C4F2", labconf: "2" },
    { codigo: "88141", color:"#D6A99A", labconf: "N" },
    { codigo: "99402.08", color:"#E16036", labconf: "2" },
    { codigo: "C0011", color:"#C8AB83", labconf: "2" },
    { codigo: "99401", color:"#85CB33", labconf: "" },
    { codigo: "96150.03", color:"#55868C", labconf: "" },
    { codigo: "99402.09", color:"#FF8E72", labconf: "" }
  ];

  valoresAdultoDr1er = [
    { codigo: "Z000", color:"#D2BF55",  labconf: "" },
    { codigo: "C8002", color:"#FFEED6", labconf: "1" },
    { codigo: "Z019", labconf: "DNT" },
    { codigo: "E46X", labconf: "IMC" },
    { codigo: "Z006", labconf: "IMC" },
    { codigo: "E660", labconf: "IMC" },
    { codigo: "E669", labconf: "IMC" },
    { codigo: "U8170", labconf: "RSM" },
    { codigo: "U8170", labconf: "RSA" },
    { codigo: "U8170", labconf: "RMA" },
    { codigo: "99199.22", labconf: "N" },
    { codigo: "99401.13", color:"#bb8588", labconf: "" },
    { codigo: "99401", color:"#85CB33", labconf: "" },
    { codigo: "Z017", labconf: "" },
    { codigo: "82947", labconf: "" },
    { codigo: "84478", labconf: "" },
    { codigo: "82465", labconf: "" },
    { codigo: "85018", labconf: "1" },
    { codigo: "81003", labconf: "" },
    { codigo: "99401", color:"#85CB33", labconf: "1" },
    { codigo: "99403.01", color:"#4CE0B3", labconf: "1" },
    { codigo: "Z128", labconf: "N" },
    { codigo: "96150.01", labconf: "" },
    { codigo: "99402.09", color:"#FF8E72", labconf: "" },
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
    { codigo: "84152", color:"#edc531", labconf: "" },
    { codigo: "99386.03", labconf: "N" },
    { codigo: "82270", color:"#a3a380", labconf: "" }
  ];
  valoresAdultoOtros2do = [
    { codigo: "88141", color:"#D6A99A", labconf: "N" },
    { codigo: "99402.08", color:"#E16036", labconf: "2" },
    { codigo: "96150.03", color:"#55868C", labconf: "" },
    { codigo: "99402.09", color:"#FF8E72", labconf: "2" },
    { codigo: "99402.03", color:"#91C4F2", labconf: "2" },
    { codigo: "99402.04", color:"#9D79BC", labconf: "2" },
    { codigo: "C0011", color:"#C8AB83", labconf: "2" },
    { codigo: "99401", color:"#85CB33", labconf: "" }
  ];
  valoresAdultoDr2do = [
    { codigo: "Z000", color:"#D2BF55",  labconf: "" },
    { codigo: "C8002", color:"#FFEED6", labconf: "TA" },
    { codigo: "99403.01", color:"#4CE0B3", labconf: "2" },
    { codigo: "U262", color:"#5465ff", labconf: "DNT" },
    { codigo: "84152", color:"#edc531", labconf: "RN" },
    { codigo: "99401.13", color:"#bb8588", labconf: "" },
    { codigo: "99401", color:"#85CB33", labconf: "2" },
    { codigo: "96150.02", color:"#FBBFCA", labconf: "2" },
    { codigo: "99402.09", color:"#FF8E72", labconf: "2" },
    { codigo: "82270", color:"#a3a380", labconf: "N" },
    { codigo: "U0041", color:"#f85e00", labconf: "" }
  ];
  valoresAdultoOtros1er = [
    { codigo: "88141", color:"#D6A99A", labconf: "N" },
    { codigo: "88141.01", color:"#ff99c8", labconf: "N" },
    { codigo: "99402.08", color:"#E16036", labconf: "1" },
    { codigo: "99402.04", color:"#9D79BC", labconf: "1" },
    { codigo: "99402.06", color:"#fcf6bd", labconf: "1" },
    { codigo: "99402.03", color:"#91C4F2", labconf: "1" },
    { codigo: "99208", color:"#8CA0D7", labconf: "" },
    { codigo: "90714", color:"#ff97b7", labconf: "1" },
    { codigo: "90658", color:"#ffb700", labconf: "" },
    { codigo: "90749.02", color:"#06d6a0", labconf: "" },
    { codigo: "C0011", color:"#C8AB83", labconf: "1" },
    { codigo: "99401", color:"#85CB33", labconf: "" },
    { codigo: "96150.04", color:"#6096ba", labconf: "" },
    { codigo: "99402.09", color:"#FF8E72", labconf: "" }
  ];

  valoresAdulMayorDr1er = [
    { codigo: "99387", color:"#c200fb", labconf: "AS" },
    { codigo: "Z636.1", color:"#6a994e", labconf: "" },
    { codigo: "C8002", color:"#FFEED6", labconf: "1" },
    { codigo: "99401", color:"#85CB33", labconf: "1" },
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
    { codigo: "99401.13", color:"#bb8588", labconf: "" },
    { codigo: "Z011", labconf: "N" },
    { codigo: "99401.13", color:"#bb8588", labconf: "1" },
    { codigo: "96150.01", labconf: "" },
    { codigo: "99402.09", color:"#FF8E72", labconf: "" },
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
    { codigo: "82270", color:"#a3a380", labconf: "" }
  ];
  valoresAdulMayorOtros1er = [
    { codigo: "96150.03", color:"#55868C", labconf: "" },
    { codigo: "99402.09", color:"#FF8E72", labconf: "" },
    { codigo: "88141", color:"#D6A99A", labconf: "" },
    { codigo: "99402.08", color:"#E16036", labconf: "1" },
    { codigo: "90658", color:"#ffb700", labconf: "" },
    { codigo: "C0011", color:"#C8AB83", labconf: "1" }
  ];
  valoresAdulMayorDr2do = [
    { codigo: "99387", color:"#c200fb", labconf: "AS" },
    { codigo: "Z636.1", color:"#6a994e", labconf: "" },
    { codigo: "C8002", color:"#FFEED6", labconf: "TA" },
    { codigo: "99401", color:"#85CB33", labconf: "2" },
    { codigo: "99403.01", color:"#4CE0B3", labconf: "2" },
    { codigo: "99401.13", color:"#bb8588", labconf: "2" },
    { codigo: "96150.07", color:"#717744", labconf: "" },
    { codigo: "99402.09", color:"#FF8E72", labconf: "2" },
    { codigo: "84152", color:"#edc531", labconf: "N" },
    { codigo: "82270", color:"#a3a380", labconf: "N" }
  ];
  valoresAdulMayorOtros2do = [
    { codigo: "88141", color:"#D6A99A", labconf: "" },
    { codigo: "C011", color:"#ccff66", labconf: "2" },
    { codigo: "96150.02", color:"#FBBFCA", labconf: "" },
    { codigo: "99402.09", color:"#FF8E72", labconf: "2" },
    { codigo: "99402.08", color:"#E16036", labconf: "2" }
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
  valoresRepeJoven = [
    { codigo: "Z000", color:"#D2BF55",  contador: 0 },
    { codigo: "C8002", color:"#FFEED6", contador: 0 },
    { codigo: "99403.01", color:"#4CE0B3", contador: 0 },
    { codigo: "99401", color:"#85CB33", contador: 0 },
    { codigo: "99173", contador: 0 },
    { codigo: "99208", color:"#8CA0D7", contador: 0 },
    { codigo: "99402.04", color:"#9D79BC", contador: 0 },
    { codigo: "99402.03", color:"#91C4F2", contador: 0 },
    { codigo: "88141", color:"#D6A99A", contador: 0 },
    { codigo: "99402.08", color:"#E16036", contador: 0 },
    { codigo: "C011", color:"#ccff66", contador: 0 },
    { codigo: "99401", color:"#85CB33", contador: 0 },
    { codigo: "99402.09", color:"#FF8E72", contador: 0 }
  ];
  valoresRepeJovenOtros = [
    { codigo: "Z000", color:"#D2BF55",  contador: 0 },
    { codigo: "C8002", color:"#FFEED6", contador: 0 },
    { codigo: "99403.01", color:"#4CE0B3", contador: 0 },
    { codigo: "99401", color:"#85CB33", contador: 0 },
    { codigo: "99173", contador: 0 },
    { codigo: "99208", color:"#8CA0D7", contador: 0 },
    { codigo: "99402.04", color:"#9D79BC", contador: 0 },
    { codigo: "99402.03", color:"#91C4F2", contador: 0 },
    { codigo: "88141", color:"#D6A99A", contador: 0 },
    { codigo: "99402.08", color:"#E16036", contador: 0 },
    { codigo: "C011", color:"#ccff66", contador: 0 },
    { codigo: "99401", color:"#85CB33", contador: 0 },
    { codigo: "99402.09", color:"#FF8E72", contador: 0 }
  ];
  compararJoven2da(codigo: any, tipo: number): string{
    let var1 = "NO";

    this.codigosRealizados.forEach((e:any) =>{
      if (e.codigo === codigo.codigo) {
        var1 = 'SI CODIGO';
        if (e.labconf === codigo.labconf) {
          var1 = 'SI CODIGO y LABCONF'
          if (codigo.labconf === ""){
            var1 = 'SI CODIGO';
          }
        }else{
          var1 = 'SI CODIGO , NO LABCONF ('+e.labconf+')';
        }
        let cod = e.codigo;
        let contador = 1;
        if(tipo==1){
          this.valoresRepeJoven.forEach((e) => {
            if(e.codigo == cod) {
              e.contador += 1
              contador = e.contador
              console.log("contador: ", e.contador);
            }
            // if(e.contador == 2) var1 += ' , SI 2 VECES';
          });
        }else{
          this.valoresRepeJovenOtros.forEach((e) => {
            if(e.codigo == cod) {
              e.contador += 1
              contador = e.contador
              console.log("contador: ", e.contador);
            }
            // if(e.contador == 2) var1 += ' , SI 2 VECES';
          });
        }
        var1 += ' , APARECE '+contador+' VECES';
      }
    })
    return var1;
  }

  valoresRepeAdulto = [
    { codigo: "Z000", color:"#D2BF55",  contador: 0 },
    { codigo: "C8002", color:"#FFEED6", contador: 0 },
    { codigo: "99403.01", color:"#4CE0B3", contador: 0 },
    { codigo: "84152", color:"#edc531", contador: 0 },
    { codigo: "99401.13", color:"#bb8588", contador: 0 },
    { codigo: "99401", color:"#85CB33", contador: 0 },
    { codigo: "99402.09", color:"#FF8E72", contador: 0 },
    { codigo: "82270", color:"#a3a380", contador: 0 },
    { codigo: "88141", color:"#D6A99A", contador: 0 },
    { codigo: "99402.08", color:"#E16036", contador: 0 },
    { codigo: "99402.09", color:"#FF8E72", contador: 0 },
    { codigo: "99402.03", color:"#91C4F2", contador: 0 },
    { codigo: "99402.04", color:"#9D79BC", contador: 0 },
    { codigo: "C0011", color:"#C8AB83", contador: 0 },
    { codigo: "99401", color:"#85CB33", contador: 0 }
  ];
  valoresRepeAdultoOtros = [
    { codigo: "Z000", color:"#D2BF55",  contador: 0 },
    { codigo: "C8002", color:"#FFEED6", contador: 0 },
    { codigo: "99403.01", color:"#4CE0B3", contador: 0 },
    { codigo: "84152", color:"#edc531", contador: 0 },
    { codigo: "99401.13", color:"#bb8588", contador: 0 },
    { codigo: "99401", color:"#85CB33", contador: 0 },
    { codigo: "99402.09", color:"#FF8E72", contador: 0 },
    { codigo: "82270", color:"#a3a380", contador: 0 },
    { codigo: "88141", color:"#D6A99A", contador: 0 },
    { codigo: "99402.08", color:"#E16036", contador: 0 },
    { codigo: "99402.09", color:"#FF8E72", contador: 0 },
    { codigo: "99402.03", color:"#91C4F2", contador: 0 },
    { codigo: "99402.04", color:"#9D79BC", contador: 0 },
    { codigo: "C0011", color:"#C8AB83", contador: 0 },
    { codigo: "99401", color:"#85CB33", contador: 0 }
  ];

  compararAdulto2da(codigo: any, tipo: number): string{
    let var1 = "NO";

    this.codigosRealizados.forEach((e:any) =>{
      if (e.codigo === codigo.codigo) {
        var1 = 'SI CODIGO';
        if (e.labconf === codigo.labconf) {
          var1 = 'SI CODIGO y LABCONF'
        }
        if (codigo.labconf !== ""){
          var1 = 'SI CODIGO , NO LABCONF ('+e.labconf+')';
        }
        let cod = e.codigo;
        let contador = 1;
        if(tipo==1){
          this.valoresRepeAdulto.forEach((e) => {
            if(e.codigo == cod) {
              e.contador += 1
              contador = e.contador
              console.log("contador: ", e.contador);
            }
            // if(e.contador == 2) var1 += ' , SI 2 VECES';
          });
        }else{
          this.valoresRepeAdultoOtros.forEach((e) => {
            if(e.codigo == cod) {
              e.contador += 1
              contador = e.contador
              console.log("contador: ", e.contador);
            }
            // if(e.contador == 2) var1 += ' , SI 2 VECES';
          });
        }
        var1 += ' , APARECE '+contador+' VECES';
      }
    })
    return var1;
  }

  valoresRepeAdultoMayor = [
    { codigo: "99387", color:"#c200fb", contador: 0 },
    { codigo: "Z636.1", color:"#6a994e", contador: 0 },
    { codigo: "C8002", color:"#FFEED6", contador: 0 },
    { codigo: "99401", color:"#85CB33", contador: 0 },
    { codigo: "99403.01", color:"#4CE0B3", contador: 0 },
    { codigo: "99401.13", color:"#bb8588", contador: 0 },
    { codigo: "99402.09", color:"#FF8E72", contador: 0 },
    { codigo: "82270", color:"#a3a380", contador: 0 },
    { codigo: "88141", color:"#D6A99A", contador: 0 },
    { codigo: "C011", color:"#ccff66", contador: 0 },
    { codigo: "99402.09", color:"#FF8E72", contador: 0 },
    { codigo: "99402.08", color:"#E16036", contador: 0 }
  ];
  valoresRepeAdultoMayorOtros = [
    { codigo: "99387", color:"#c200fb", contador: 0 },
    { codigo: "Z636.1", color:"#6a994e", contador: 0 },
    { codigo: "C8002", color:"#FFEED6", contador: 0 },
    { codigo: "99401", color:"#85CB33", contador: 0 },
    { codigo: "99403.01", color:"#4CE0B3", contador: 0 },
    { codigo: "99401.13", color:"#bb8588", contador: 0 },
    { codigo: "99402.09", color:"#FF8E72", contador: 0 },
    { codigo: "82270", color:"#a3a380", contador: 0 },
    { codigo: "88141", color:"#D6A99A", contador: 0 },
    { codigo: "C011", color:"#ccff66", contador: 0 },
    { codigo: "99402.09", color:"#FF8E72", contador: 0 },
    { codigo: "99402.08", color:"#E16036", contador: 0 }
  ];
  compararAdulMayor2da(codigo: any, tipo: number): string{
    let var1 = "NO";

    this.codigosRealizados.forEach((e:any) =>{
      if (e.codigo === codigo.codigo) {
        var1 = 'SI CODIGO';
        if (e.labconf === codigo.labconf) {
          var1 = 'SI CODIGO y LABCONF'
        }
        if (codigo.labconf !== ""){
          var1 = 'SI CODIGO , NO LABCONF ('+e.labconf+')';
        }
        let cod = e.codigo;
        let contador = 1;
        if(tipo == 1){
          this.valoresRepeAdultoMayor.forEach((e) => {
            if(e.codigo == cod) {
              e.contador += 1
              contador = e.contador
              console.log("contador: ", e.contador);
            }
            // if(e.contador == 2) var1 += ' , SI 2 VECES';
          });
        }else{
          this.valoresRepeAdultoMayorOtros.forEach((e) => {
            if(e.codigo == cod) {
              e.contador += 1
              contador = e.contador
              console.log("contador: ", e.contador);
            }
            // if(e.contador == 2) var1 += ' , SI 2 VECES';
          });
        }
        var1 += ' , APARECE '+contador+' VECES';
      }
    })
    return var1;
  }

  isUndefined(val: any): string{
    return (val == undefined ? "" : String(val))
  }

}
