import { Component, OnInit, VERSION } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent implements OnInit {
  name = 'Angular ' + VERSION.major;
  xml: any = [];
  start: string;
  voucherEnd: string;
  end: string;
  totalRecords: number;
  type: number[] = [0, 5, 12, 18];
  mastersStart: string;
  mastersEnd: string;
  mastersEntry: string;
  mastersXML: any = [];
  stateList = [
    {
      stateCode: '01',
      stateName: 'Jammu & Kashmir',
    },
    {
      stateCode: '02',
      stateName: 'Himachal Pradesh',
    },
    {
      stateCode: '03',
      stateName: 'Punjab',
    },
    {
      stateCode: '04',
      stateName: 'Chandigarh',
    },
    {
      stateCode: '05',
      stateName: 'Uttarakhand',
    },
    {
      stateCode: '06',
      stateName: 'Haryana',
    },
    {
      stateCode: '07',
      stateName: 'Delhi',
    },
    {
      stateCode: '08',
      stateName: 'Rajasthan',
    },
    {
      stateCode: '09',
      stateName: 'Uttar Pradesh',
    },
    {
      stateCode: '10',
      stateName: 'Bihar',
    },
    {
      stateCode: '11',
      stateName: 'Sikkim',
    },
    {
      stateCode: '12',
      stateName: 'Arunachal Pradesh',
    },
    {
      stateCode: '13',
      stateName: 'Nagaland',
    },
    {
      stateCode: '14',
      stateName: 'Manipur',
    },
    {
      stateCode: '15',
      stateName: 'Mizoram',
    },
    {
      stateCode: '16',
      stateName: 'Tripura',
    },
    {
      stateCode: '17',
      stateName: 'Meghalaya',
    },
    {
      stateCode: '18',
      stateName: 'Assam',
    },
    {
      stateCode: '19',
      stateName: 'West Bengal',
    },
    {
      stateCode: '20',
      stateName: 'Jharkhand',
    },
    {
      stateCode: '21',
      stateName: 'Orissa',
    },
    {
      stateCode: '22',
      stateName: 'Chhattisgarh',
    },
    {
      stateCode: '23',
      stateName: 'Madhya Pradesh',
    },
    {
      stateCode: '24',
      stateName: 'Gujarat',
    },
    {
      stateCode: '25',
      stateName: 'Daman & Diu',
    },
    {
      stateCode: '26',
      stateName: 'Dadra & Nagar Haveli',
    },
    {
      stateCode: '27',
      stateName: 'Maharashtra',
    },
    {
      stateCode: '28',
      stateName: 'Andhra Pradesh',
    },
    {
      stateCode: '29',
      stateName: 'Karnataka',
    },
    {
      stateCode: '30',
      stateName: 'Goa',
    },
    {
      stateCode: '31',
      stateName: 'Lakshadweep',
    },
    {
      stateCode: '32',
      stateName: 'Kerala',
    },
    {
      stateCode: '33',
      stateName: 'Tamil Nadu',
    },
    {
      stateCode: '34',
      stateName: 'Puducherry',
    },
    {
      stateCode: '35',
      stateName: 'Andaman & Nicobar Islands',
    },
    {
      stateCode: '36',
      stateName: 'Telengana',
    },
    {
      stateCode: '37',
      stateName: 'Andhra Pradesh',
    },
  ];

  ngOnInit() {
    this.initialize();
    // const str = 'ATHARV MEDICAL & SURGICALS KANTH';
    // console.log(str.replace('&', '&amp;'));
  }

  prepareData(data: any) {
    this.xml.push(this.start);

    for (let i = 0; i < data.length; i++) {
      debugger;
      const item = data[i];
      this.prepareVoucher(item);
      this.ledgerEntries(
        item.partyName.trim().replace('&', '&amp;'),
        -item.total
      );

      // this loop for only ledger entries
      for (let j = 0; j < this.type.length; j++) {
        const colName = `SALE ${
          (item.gstNo.trim().substring(0, 2) === '09' ? '' : 'EX UP ') +
          this.type[j]
        }% (T.I.)`;

        const tax = `TAX ${
          (item.gstNo.trim().substring(0, 2) === '09' ? '' : 'EX UP ') +
          this.type[j]
        }%`;

        if (item[colName] && item.gstNo.trim().substring(0, 2) === '09') {
          if (this.type[j] === 0) {
            this.ledgerEntries('SALE EXEMPTED (T.I.)', item[colName], true);
          } else {
            this.ledgerEntries(colName, item[colName], true);
            const taxPercent = this.type[j] / 2 + '%';
            const taxValue = item[tax];
            this.ledgerEntries('CGST ' + taxPercent, taxValue, true);
            this.ledgerEntries('SGST ' + taxPercent, taxValue, true);
          }
        } else if (item[colName] && item.gstNo.trim().substring(0, 2) != '09') {
          if (this.type[j] === 0) {
            this.ledgerEntries(
              'SALE EX UP EXEMPTED (T.I.)',
              item[colName],
              true
            );
          } else {
            this.ledgerEntries(colName, item[colName], true);
            const taxPercent = this.type[j] + '%';
            const taxValue = item[tax];
            this.ledgerEntries('IGST ' + taxPercent, taxValue, true);
          }
        }
      }
      this.ledgerEntries('R/O', item['RO'], true);
      this.xml.push(this.voucherEnd);
    }
    this.xml.push(this.end);
  }

  prepareVoucher(data: any) {
    const entry = `<TALLYMESSAGE xmlns:UDF="TallyUDF">
    <VOUCHER VCHTYPE="Sales"
             ACTION="Create"
             OBJVIEW="Invoice Voucher View">
      <OLDAUDITENTRYIDS.LIST TYPE="Number">
        <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
      </OLDAUDITENTRYIDS.LIST>
      <DATE>${data.billDate}</DATE>
      <GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>
      <VATDEALERTYPE>Regular</VATDEALERTYPE>
      <STATENAME>${this.getStateName(data.gstNo)}</STATENAME>
      <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
      <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
      <PARTYGSTIN>${data.gstNo}</PARTYGSTIN>
      <PLACEOFSUPPLY>${this.getStateName(data.gstNo)}</PLACEOFSUPPLY>
      <PARTYNAME>${data.partyName}</PARTYNAME>
      <PARTYLEDGERNAME>${data.partyName}</PARTYLEDGERNAME>
      <PARTYMAILINGNAME>${data.partyName}</PARTYMAILINGNAME>
      <CONSIGNEEGSTIN>${data.gstNo}</CONSIGNEEGSTIN>
      <CONSIGNEEMAILINGNAME>${data.partyName}</CONSIGNEEMAILINGNAME>
      <CONSIGNEESTATENAME>${this.getStateName(data.gstNo)}</CONSIGNEESTATENAME>
      <VOUCHERNUMBER>${data.billNo}</VOUCHERNUMBER>
      <BASICBASEPARTYNAME>${data.partyName}</BASICBASEPARTYNAME>
      <CSTFORMISSUETYPE/>
      <CSTFORMRECVTYPE/>
      <FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>
      <PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
      <BASICBUYERNAME>${data.partyName}</BASICBUYERNAME>
      <CONSIGNEECOUNTRYNAME>India</CONSIGNEECOUNTRYNAME>
      <VCHGSTCLASS/>
      <VCHENTRYMODE>Accounting Invoice</VCHENTRYMODE>
      <DIFFACTUALQTY>No</DIFFACTUALQTY>
      <ISMSTFROMSYNC>No</ISMSTFROMSYNC>
      <ISDELETED>No</ISDELETED>
      <ISSECURITYONWHENENTERED>No</ISSECURITYONWHENENTERED>
      <ASORIGINAL>No</ASORIGINAL>
      <AUDITED>No</AUDITED>
      <FORJOBCOSTING>No</FORJOBCOSTING>
      <ISOPTIONAL>No</ISOPTIONAL>
      <EFFECTIVEDATE>${data.billDate}</EFFECTIVEDATE>
      <USEFOREXCISE>No</USEFOREXCISE>
      <ISFORJOBWORKIN>No</ISFORJOBWORKIN>
      <ALLOWCONSUMPTION>No</ALLOWCONSUMPTION>
      <USEFORINTEREST>No</USEFORINTEREST>
      <USEFORGAINLOSS>No</USEFORGAINLOSS>
      <USEFORGODOWNTRANSFER>No</USEFORGODOWNTRANSFER>
      <USEFORCOMPOUND>No</USEFORCOMPOUND>
      <USEFORSERVICETAX>No</USEFORSERVICETAX>
      <ISONHOLD>No</ISONHOLD>
      <ISBOENOTAPPLICABLE>No</ISBOENOTAPPLICABLE>
      <ISGSTSECSEVENAPPLICABLE>No</ISGSTSECSEVENAPPLICABLE>
      <ISEXCISEVOUCHER>No</ISEXCISEVOUCHER>
      <EXCISETAXOVERRIDE>No</EXCISETAXOVERRIDE>
      <USEFORTAXUNITTRANSFER>No</USEFORTAXUNITTRANSFER>
      <IGNOREPOSVALIDATION>No</IGNOREPOSVALIDATION>
      <EXCISEOPENING>No</EXCISEOPENING>
      <USEFORFINALPRODUCTION>No</USEFORFINALPRODUCTION>
      <ISTDSOVERRIDDEN>No</ISTDSOVERRIDDEN>
      <ISTCSOVERRIDDEN>No</ISTCSOVERRIDDEN>
      <ISTDSTCSCASHVCH>No</ISTDSTCSCASHVCH>
      <INCLUDEADVPYMTVCH>No</INCLUDEADVPYMTVCH>
      <ISSUBWORKSCONTRACT>No</ISSUBWORKSCONTRACT>
      <ISVATOVERRIDDEN>No</ISVATOVERRIDDEN>
      <IGNOREORIGVCHDATE>No</IGNOREORIGVCHDATE>
      <ISVATPAIDATCUSTOMS>No</ISVATPAIDATCUSTOMS>
      <ISDECLAREDTOCUSTOMS>No</ISDECLAREDTOCUSTOMS>
      <ISSERVICETAXOVERRIDDEN>No</ISSERVICETAXOVERRIDDEN>
      <ISISDVOUCHER>No</ISISDVOUCHER>
      <ISEXCISEOVERRIDDEN>No</ISEXCISEOVERRIDDEN>
      <ISEXCISESUPPLYVCH>No</ISEXCISESUPPLYVCH>
      <ISGSTOVERRIDDEN>No</ISGSTOVERRIDDEN>
      <GSTNOTEXPORTED>No</GSTNOTEXPORTED>
      <IGNOREGSTINVALIDATION>No</IGNOREGSTINVALIDATION>
      <ISGSTREFUND>No</ISGSTREFUND>
      <OVRDNEWAYBILLAPPLICABILITY>No</OVRDNEWAYBILLAPPLICABILITY>
      <ISVATPRINCIPALACCOUNT>No</ISVATPRINCIPALACCOUNT>
      <IGNOREEINVVALIDATION>No</IGNOREEINVVALIDATION>
      <IRNJSONEXPORTED>No</IRNJSONEXPORTED>
      <IRNCANCELLED>No</IRNCANCELLED>
      <ISSHIPPINGWITHINSTATE>No</ISSHIPPINGWITHINSTATE>
      <ISOVERSEASTOURISTTRANS>No</ISOVERSEASTOURISTTRANS>
      <ISDESIGNATEDZONEPARTY>No</ISDESIGNATEDZONEPARTY>
      <ISCANCELLED>No</ISCANCELLED>
      <HASCASHFLOW>No</HASCASHFLOW>
      <ISPOSTDATED>No</ISPOSTDATED>
      <USETRACKINGNUMBER>No</USETRACKINGNUMBER>
      <ISINVOICE>Yes</ISINVOICE>
      <MFGJOURNAL>No</MFGJOURNAL>
      <HASDISCOUNTS>No</HASDISCOUNTS>
      <ASPAYSLIP>No</ASPAYSLIP>
      <ISCOSTCENTRE>No</ISCOSTCENTRE>
      <ISSTXNONREALIZEDVCH>No</ISSTXNONREALIZEDVCH>
      <ISEXCISEMANUFACTURERON>No</ISEXCISEMANUFACTURERON>
      <ISBLANKCHEQUE>No</ISBLANKCHEQUE>
      <ISVOID>No</ISVOID>
      <ORDERLINESTATUS>No</ORDERLINESTATUS>
      <VATISAGNSTCANCSALES>No</VATISAGNSTCANCSALES>
      <VATISPURCEXEMPTED>No</VATISPURCEXEMPTED>
      <ISVATRESTAXINVOICE>No</ISVATRESTAXINVOICE>
      <VATISASSESABLECALCVCH>No</VATISASSESABLECALCVCH>
      <ISVATDUTYPAID>Yes</ISVATDUTYPAID>
      <ISDELIVERYSAMEASCONSIGNEE>No</ISDELIVERYSAMEASCONSIGNEE>
      <ISDISPATCHSAMEASCONSIGNOR>No</ISDISPATCHSAMEASCONSIGNOR>
      <ISDELETEDVCHRETAINED>No</ISDELETEDVCHRETAINED>
      <CHANGEVCHMODE>No</CHANGEVCHMODE>
      <RESETIRNQRCODE>No</RESETIRNQRCODE>
      <EWAYBILLDETAILS.LIST>      </EWAYBILLDETAILS.LIST>
      <EXCLUDEDTAXATIONS.LIST>      </EXCLUDEDTAXATIONS.LIST>
      <OLDAUDITENTRIES.LIST>      </OLDAUDITENTRIES.LIST>
      <ACCOUNTAUDITENTRIES.LIST>      </ACCOUNTAUDITENTRIES.LIST>
      <AUDITENTRIES.LIST>      </AUDITENTRIES.LIST>
      <DUTYHEADDETAILS.LIST>      </DUTYHEADDETAILS.LIST>
      <ALLINVENTORYENTRIES.LIST>      </ALLINVENTORYENTRIES.LIST>
      <SUPPLEMENTARYDUTYHEADDETAILS.LIST>      </SUPPLEMENTARYDUTYHEADDETAILS.LIST>
      <EWAYBILLERRORLIST.LIST>      </EWAYBILLERRORLIST.LIST>
      <IRNERRORLIST.LIST>      </IRNERRORLIST.LIST>
      <INVOICEDELNOTES.LIST>      </INVOICEDELNOTES.LIST>
      <INVOICEORDERLIST.LIST>      </INVOICEORDERLIST.LIST>
      <INVOICEINDENTLIST.LIST>      </INVOICEINDENTLIST.LIST>
      <ATTENDANCEENTRIES.LIST>      </ATTENDANCEENTRIES.LIST>
      <ORIGINVOICEDETAILS.LIST>      </ORIGINVOICEDETAILS.LIST>
      <INVOICEEXPORTLIST.LIST>      </INVOICEEXPORTLIST.LIST>`;
    this.xml.push(entry);
  }

  ledgerEntries(desc: string, amt: number, flag: boolean = false) {
    const entry = `
    <LEDGERENTRIES.LIST>
     <OLDAUDITENTRYIDS.LIST TYPE='Number'>
      <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
     </OLDAUDITENTRYIDS.LIST>
     <LEDGERNAME>${desc}</LEDGERNAME>
     <GSTCLASS/>
     <ISDEEMEDPOSITIVE>${!flag ? 'Yes' : 'No'}</ISDEEMEDPOSITIVE>
     <LEDGERFROMITEM>No</LEDGERFROMITEM>
     <REMOVEZEROENTRIES>No</REMOVEZEROENTRIES>
     <ISPARTYLEDGER>${flag ? 'Yes' : 'No'}</ISPARTYLEDGER>
     <ISLASTDEEMEDPOSITIVE>${!flag ? 'Yes' : 'No'}</ISLASTDEEMEDPOSITIVE>
     <ISCAPVATTAXALTERED>No</ISCAPVATTAXALTERED>
     <ISCAPVATNOTCLAIMED>No</ISCAPVATNOTCLAIMED>
     <AMOUNT>${amt}</AMOUNT>
      ${flag ? '' : '<VATEXPAMOUNT>' + amt + '</VATEXPAMOUNT>'}
     <SERVICETAXDETAILS.LIST>       </SERVICETAXDETAILS.LIST>
     <BANKALLOCATIONS.LIST>       </BANKALLOCATIONS.LIST>
     <BILLALLOCATIONS.LIST>       </BILLALLOCATIONS.LIST>
       <INTERESTCOLLECTION.LIST>       </INTERESTCOLLECTION.LIST>
     <OLDAUDITENTRIES.LIST>       </OLDAUDITENTRIES.LIST>
       <ACCOUNTAUDITENTRIES.LIST>       </ACCOUNTAUDITENTRIES.LIST>
     <AUDITENTRIES.LIST>       </AUDITENTRIES.LIST>
     <INPUTCRALLOCS.LIST>       </INPUTCRALLOCS.LIST>
     <DUTYHEADDETAILS.LIST>       </DUTYHEADDETAILS.LIST>
       <EXCISEDUTYHEADDETAILS.LIST>       </EXCISEDUTYHEADDETAILS.LIST>
     <RATEDETAILS.LIST>       </RATEDETAILS.LIST>
     <SUMMARYALLOCS.LIST>       </SUMMARYALLOCS.LIST>
     <STPYMTDETAILS.LIST>       </STPYMTDETAILS.LIST>
       <EXCISEPAYMENTALLOCATIONS.LIST>       </EXCISEPAYMENTALLOCATIONS.LIST>
       <TAXBILLALLOCATIONS.LIST>       </TAXBILLALLOCATIONS.LIST>
       <TAXOBJECTALLOCATIONS.LIST>       </TAXOBJECTALLOCATIONS.LIST>
       <TDSEXPENSEALLOCATIONS.LIST>       </TDSEXPENSEALLOCATIONS.LIST>
       <COSTTRACKALLOCATIONS.LIST>       </COSTTRACKALLOCATIONS.LIST>
     <REFVOUCHERDETAILS.LIST>       </REFVOUCHERDETAILS.LIST>
       <INVOICEWISEDETAILS.LIST>       </INVOICEWISEDETAILS.LIST>
     <VATITCDETAILS.LIST>       </VATITCDETAILS.LIST>
     <ADVANCETAXDETAILS.LIST>       </ADVANCETAXDETAILS.LIST>
    </LEDGERENTRIES.LIST>
   `;
    this.xml.push(entry);
  }

  initialize() {
    this.start = `<ENVELOPE>
    <HEADER>
      <TALLYREQUEST>Import Data</TALLYREQUEST>
    </HEADER>
    <BODY>
      <IMPORTDATA>
        <REQUESTDESC>
          <REPORTNAME>All Masters</REPORTNAME>
          <STATICVARIABLES>
            <SVCURRENTCOMPANY>DEEP PHARMA TRADERS</SVCURRENTCOMPANY>
          </STATICVARIABLES>
        </REQUESTDESC>
        <REQUESTDATA>`;

    this.voucherEnd = `<PAYROLLMODEOFPAYMENT.LIST>      </PAYROLLMODEOFPAYMENT.LIST>
    <ATTDRECORDS.LIST>      </ATTDRECORDS.LIST>
    <GSTEWAYCONSIGNORADDRESS.LIST>      </GSTEWAYCONSIGNORADDRESS.LIST>
    <GSTEWAYCONSIGNEEADDRESS.LIST>      </GSTEWAYCONSIGNEEADDRESS.LIST>
    <TEMPGSTRATEDETAILS.LIST>      </TEMPGSTRATEDETAILS.LIST>
    </VOUCHER>
    </TALLYMESSAGE>`;

    this.end = `<TALLYMESSAGE xmlns:UDF="TallyUDF">
    <COMPANY>
      <REMOTECMPINFO.LIST MERGE="Yes">
        <NAME>732743c1-09c0-4e28-b78d-832f56733720</NAME>
        <REMOTECMPNAME>DEEP PHARMA TRADERS</REMOTECMPNAME>
        <REMOTECMPSTATE>Uttar Pradesh</REMOTECMPSTATE>
      </REMOTECMPINFO.LIST>
    </COMPANY>
      </TALLYMESSAGE>
    </REQUESTDATA>
    </IMPORTDATA>
    </BODY>
    </ENVELOPE>`;
  }

  getStateName(gstNo: string) {
    let item = this.stateList.find(
      (t) => t.stateCode === gstNo.trim().substring(0, 2)
    );
    return item.stateName;
  }

  prepareMasters(masters: any) {
    this.mastersStart = `<ENVELOPE>
    <HEADER>
      <TALLYREQUEST>Import Data</TALLYREQUEST>
    </HEADER>
    <BODY>
      <IMPORTDATA>
        <REQUESTDESC>
          <REPORTNAME>All Masters</REPORTNAME>
          <STATICVARIABLES>
            <SVCURRENTCOMPANY>DEEP PHARMA TRADERS</SVCURRENTCOMPANY>
          </STATICVARIABLES>
        </REQUESTDESC>
        <REQUESTDATA>`;
    this.mastersXML.push(this.mastersStart);

    masters.forEach((el) => {
      this.prepareMastersList(el);
    });

    this.mastersEnd = `</REQUESTDATA>
        </IMPORTDATA>
      </BODY>
    </ENVELOPE>`;
    this.mastersXML.push(this.mastersEnd);
  }

  prepareMastersList(data: any) {
    this.mastersEntry = `<TALLYMESSAGE xmlns:UDF="TallyUDF">
        <LEDGER NAME="${data.partyName
          .trim()
          .replace('&', '&amp;')}" RESERVEDNAME="">
        <MAILINGNAME.LIST TYPE="String">
          <MAILINGNAME>${data.partyName
            .trim()
            .replace('&', '&amp;')}</MAILINGNAME>
        </MAILINGNAME.LIST>
        <OLDAUDITENTRYIDS.LIST TYPE="Number">
          <OLDAUDITENTRYIDS>-1</OLDAUDITENTRYIDS>
        </OLDAUDITENTRYIDS.LIST>
        <GUID>732743c1-09c0-4e28-b78d-832f56733720-000000b9</GUID>
        <CURRENCYNAME>₹</CURRENCYNAME>
        <PRIORSTATENAME>${this.getStateName(data.gstNo)}</PRIORSTATENAME>
        <COUNTRYNAME>India</COUNTRYNAME>
        <GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>
        <VATDEALERTYPE>Regular</VATDEALERTYPE>
        <PARENT>Sundry Debtors</PARENT>
        <TAXCLASSIFICATIONNAME/>
        <TAXTYPE>Others</TAXTYPE>
        <COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
        <GSTTYPE/>
        <APPROPRIATEFOR/>
        <PARTYGSTIN>${data.gstNo}</PARTYGSTIN>
        <LEDSTATENAME>${this.getStateName(data.gstNo)}</LEDSTATENAME>
        <SERVICECATEGORY/>
        <EXCISELEDGERCLASSIFICATION/>
        <EXCISEDUTYTYPE/>
        <EXCISENATUREOFPURCHASE/>
        <LEDGERFBTCATEGORY/>
        <ISBILLWISEON>No</ISBILLWISEON>
        <ISCOSTCENTRESON>No</ISCOSTCENTRESON>
        <ISINTERESTON>No</ISINTERESTON>
        <ALLOWINMOBILE>No</ALLOWINMOBILE>
        <ISCOSTTRACKINGON>No</ISCOSTTRACKINGON>
        <ISBENEFICIARYCODEON>No</ISBENEFICIARYCODEON>
        <ISEXPORTONVCHCREATE>No</ISEXPORTONVCHCREATE>
        <PLASINCOMEEXPENSE>No</PLASINCOMEEXPENSE>
        <ISUPDATINGTARGETID>No</ISUPDATINGTARGETID>
        <ISDELETED>No</ISDELETED>
        <ISSECURITYONWHENENTERED>No</ISSECURITYONWHENENTERED>
        <ASORIGINAL>Yes</ASORIGINAL>
        <ISCONDENSED>No</ISCONDENSED>
        <AFFECTSSTOCK>No</AFFECTSSTOCK>
        <ISRATEINCLUSIVEVAT>No</ISRATEINCLUSIVEVAT>
        <FORPAYROLL>No</FORPAYROLL>
        <ISABCENABLED>No</ISABCENABLED>
        <ISCREDITDAYSCHKON>No</ISCREDITDAYSCHKON>
        <INTERESTONBILLWISE>No</INTERESTONBILLWISE>
        <OVERRIDEINTEREST>No</OVERRIDEINTEREST>
        <OVERRIDEADVINTEREST>No</OVERRIDEADVINTEREST>
        <USEFORVAT>No</USEFORVAT>
        <IGNORETDSEXEMPT>No</IGNORETDSEXEMPT>
        <ISTCSAPPLICABLE>No</ISTCSAPPLICABLE>
        <ISTDSAPPLICABLE>No</ISTDSAPPLICABLE>
        <ISFBTAPPLICABLE>No</ISFBTAPPLICABLE>
        <ISGSTAPPLICABLE>No</ISGSTAPPLICABLE>
        <ISEXCISEAPPLICABLE>No</ISEXCISEAPPLICABLE>
        <ISTDSEXPENSE>No</ISTDSEXPENSE>
        <ISEDLIAPPLICABLE>No</ISEDLIAPPLICABLE>
        <ISRELATEDPARTY>No</ISRELATEDPARTY>
        <USEFORESIELIGIBILITY>No</USEFORESIELIGIBILITY>
        <ISINTERESTINCLLASTDAY>No</ISINTERESTINCLLASTDAY>
        <APPROPRIATETAXVALUE>No</APPROPRIATETAXVALUE>
        <ISBEHAVEASDUTY>No</ISBEHAVEASDUTY>
        <INTERESTINCLDAYOFADDITION>No</INTERESTINCLDAYOFADDITION>
        <INTERESTINCLDAYOFDEDUCTION>No</INTERESTINCLDAYOFDEDUCTION>
        <ISOTHTERRITORYASSESSEE>No</ISOTHTERRITORYASSESSEE>
        <IGNOREMISMATCHWITHWARNING>No</IGNOREMISMATCHWITHWARNING>
        <USEASNOTIONALBANK>No</USEASNOTIONALBANK>
        <OVERRIDECREDITLIMIT>No</OVERRIDECREDITLIMIT>
        <ISAGAINSTFORMC>No</ISAGAINSTFORMC>
        <ISCHEQUEPRINTINGENABLED>Yes</ISCHEQUEPRINTINGENABLED>
        <ISPAYUPLOAD>No</ISPAYUPLOAD>
        <ISPAYBATCHONLYSAL>No</ISPAYBATCHONLYSAL>
        <ISBNFCODESUPPORTED>No</ISBNFCODESUPPORTED>
        <ALLOWEXPORTWITHERRORS>No</ALLOWEXPORTWITHERRORS>
        <CONSIDERPURCHASEFOREXPORT>No</CONSIDERPURCHASEFOREXPORT>
        <ISTRANSPORTER>No</ISTRANSPORTER>
        <USEFORNOTIONALITC>No</USEFORNOTIONALITC>
        <ISECOMMOPERATOR>No</ISECOMMOPERATOR>
        <OVERRIDEBASEDONREALIZATION>No</OVERRIDEBASEDONREALIZATION>
        <SHOWINPAYSLIP>No</SHOWINPAYSLIP>
        <USEFORGRATUITY>No</USEFORGRATUITY>
        <ISTDSPROJECTED>No</ISTDSPROJECTED>
        <FORSERVICETAX>No</FORSERVICETAX>
        <ISINPUTCREDIT>No</ISINPUTCREDIT>
        <ISEXEMPTED>No</ISEXEMPTED>
        <ISABATEMENTAPPLICABLE>No</ISABATEMENTAPPLICABLE>
        <ISSTXPARTY>No</ISSTXPARTY>
        <ISSTXNONREALIZEDTYPE>No</ISSTXNONREALIZEDTYPE>
        <ISUSEDFORCVD>No</ISUSEDFORCVD>
        <LEDBELONGSTONONTAXABLE>No</LEDBELONGSTONONTAXABLE>
        <ISEXCISEMERCHANTEXPORTER>No</ISEXCISEMERCHANTEXPORTER>
        <ISPARTYEXEMPTED>No</ISPARTYEXEMPTED>
        <ISSEZPARTY>No</ISSEZPARTY>
        <TDSDEDUCTEEISSPECIALRATE>No</TDSDEDUCTEEISSPECIALRATE>
        <ISECHEQUESUPPORTED>No</ISECHEQUESUPPORTED>
        <ISEDDSUPPORTED>No</ISEDDSUPPORTED>
        <HASECHEQUEDELIVERYMODE>No</HASECHEQUEDELIVERYMODE>
        <HASECHEQUEDELIVERYTO>No</HASECHEQUEDELIVERYTO>
        <HASECHEQUEPRINTLOCATION>No</HASECHEQUEPRINTLOCATION>
        <HASECHEQUEPAYABLELOCATION>No</HASECHEQUEPAYABLELOCATION>
        <HASECHEQUEBANKLOCATION>No</HASECHEQUEBANKLOCATION>
        <HASEDDDELIVERYMODE>No</HASEDDDELIVERYMODE>
        <HASEDDDELIVERYTO>No</HASEDDDELIVERYTO>
        <HASEDDPRINTLOCATION>No</HASEDDPRINTLOCATION>
        <HASEDDPAYABLELOCATION>No</HASEDDPAYABLELOCATION>
        <HASEDDBANKLOCATION>No</HASEDDBANKLOCATION>
        <ISEBANKINGENABLED>No</ISEBANKINGENABLED>
        <ISEXPORTFILEENCRYPTED>No</ISEXPORTFILEENCRYPTED>
        <ISBATCHENABLED>No</ISBATCHENABLED>
        <ISPRODUCTCODEBASED>No</ISPRODUCTCODEBASED>
        <HASEDDCITY>No</HASEDDCITY>
        <HASECHEQUECITY>No</HASECHEQUECITY>
        <ISFILENAMEFORMATSUPPORTED>No</ISFILENAMEFORMATSUPPORTED>
        <HASCLIENTCODE>No</HASCLIENTCODE>
        <PAYINSISBATCHAPPLICABLE>No</PAYINSISBATCHAPPLICABLE>
        <PAYINSISFILENUMAPP>No</PAYINSISFILENUMAPP>
        <ISSALARYTRANSGROUPEDFORBRS>No</ISSALARYTRANSGROUPEDFORBRS>
        <ISEBANKINGSUPPORTED>No</ISEBANKINGSUPPORTED>
        <ISSCBUAE>No</ISSCBUAE>
        <ISBANKSTATUSAPP>No</ISBANKSTATUSAPP>
        <ISSALARYGROUPED>No</ISSALARYGROUPED>
        <USEFORPURCHASETAX>No</USEFORPURCHASETAX>
        <AUDITED>No</AUDITED>
        <SERVICETAXDETAILS.LIST/>
        <LBTREGNDETAILS.LIST/>
        <VATDETAILS.LIST/>
        <SALESTAXCESSDETAILS.LIST/>
        <GSTDETAILS.LIST/>
        <LANGUAGENAME.LIST>
          <NAME.LIST TYPE="String">
            <NAME>${data.partyName.trim().replace('&', '&amp;')}</NAME>
          </NAME.LIST>
          <LANGUAGEID>1033</LANGUAGEID>
        </LANGUAGENAME.LIST>
        <XBRLDETAIL.LIST/>
        <AUDITDETAILS.LIST/>
        <SCHVIDETAILS.LIST/>
        <EXCISETARIFFDETAILS.LIST/>
        <TCSCATEGORYDETAILS.LIST/>
        <TDSCATEGORYDETAILS.LIST/>
        <SLABPERIOD.LIST/>
        <GRATUITYPERIOD.LIST/>
        <ADDITIONALCOMPUTATIONS.LIST/>
        <EXCISEJURISDICTIONDETAILS.LIST/>
        <EXCLUDEDTAXATIONS.LIST/>
        <BANKALLOCATIONS.LIST/>
        <PAYMENTDETAILS.LIST/>
        <BANKEXPORTFORMATS.LIST/>
        <BILLALLOCATIONS.LIST/>
        <INTERESTCOLLECTION.LIST/>
        <LEDGERCLOSINGVALUES.LIST/>
        <LEDGERAUDITCLASS.LIST/>
        <OLDAUDITENTRIES.LIST/>
        <TDSEXEMPTIONRULES.LIST/>
        <DEDUCTINSAMEVCHRULES.LIST/>
        <LOWERDEDUCTION.LIST/>
        <STXABATEMENTDETAILS.LIST/>
        <LEDMULTIADDRESSLIST.LIST/>
        <STXTAXDETAILS.LIST/>
        <CHEQUERANGE.LIST/>
        <DEFAULTVCHCHEQUEDETAILS.LIST/>
        <ACCOUNTAUDITENTRIES.LIST/>
        <AUDITENTRIES.LIST/>
        <BRSIMPORTEDINFO.LIST/>
        <AUTOBRSCONFIGS.LIST/>
        <BANKURENTRIES.LIST/>
        <DEFAULTCHEQUEDETAILS.LIST/>
        <DEFAULTOPENINGCHEQUEDETAILS.LIST/>
        <CANCELLEDPAYALLOCATIONS.LIST/>
        <ECHEQUEPRINTLOCATION.LIST/>
        <ECHEQUEPAYABLELOCATION.LIST/>
        <EDDPRINTLOCATION.LIST/>
        <EDDPAYABLELOCATION.LIST/>
        <AVAILABLETRANSACTIONTYPES.LIST/>
        <LEDPAYINSCONFIGS.LIST/>
        <TYPECODEDETAILS.LIST/>
        <FIELDVALIDATIONDETAILS.LIST/>
        <INPUTCRALLOCS.LIST/>
        <TCSMETHODOFCALCULATION.LIST/>
        <GSTCLASSFNIGSTRATES.LIST/>
        <EXTARIFFDUTYHEADDETAILS.LIST/>
        <VOUCHERTYPEPRODUCTCODES.LIST/>
      </LEDGER>
      </TALLYMESSAGE>`;

    this.mastersXML.push(this.mastersEntry);
  }

  onFileChange(ev) {
    let workBook = null;
    let jsonData = null;
    const target: DataTransfer = <DataTransfer>ev.target;
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();

    reader.onload = (event) => {
      const data = reader.result;
      workBook = XLSX.read(data, { type: 'binary' });
      jsonData = workBook.SheetNames.reduce((initial, name) => {
        const sheet = workBook.Sheets[name];
        initial[name] = XLSX.utils.sheet_to_json(sheet, {
          defval: '',
        });
        return initial;
      }, {});
      console.log(jsonData.sheet1);
      // this.prepareMasters(jsonData.sheet1);
      this.prepareData(jsonData.sheet1);
      this.totalRecords = jsonData.sheet1.length;
      // console.log(this.xml);
    };
    reader.readAsBinaryString(target.files[0]);
  }

  download() {
    // this.downloadFile('masters.xml', this.mastersXML.join(''));
    this.downloadFile('report.xml', this.xml.join(''));
  }

  downloadFile(filename, textInput) {
    var element = document.createElement('a');
    element.setAttribute(
      'href',
      'data:text/plain;charset=utf-8, ' + encodeURIComponent(textInput)
    );
    element.setAttribute('download', filename);
    document.body.appendChild(element);
    element.click();
  }
}
