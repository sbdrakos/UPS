*UPPROCS.PRG PROCEDURES FOR UPSYSTEM
#include "inkey.ch"
#include "appevent.ch"
#include "gra.ch"
#include "xbp.ch"
#include "common.ch"
#include "set.ch"
#include 'xwtone.ch'
#include 'memoedit.ch'
#include "upsystnew.ch"
#include "asxml.ch"
#include 'memoedit.ch'
#include 'dbstruct.ch'
#include "dbfdbe.ch"
#include "fileio.ch"
#include "simpleio.ch"
#include "dmlb.ch"
#include "dll.ch"
#include 'ntxdbe.ch'

#include 'dcdialog.ch'
#include "dcbitmap.ch"
#include "dcgra.ch"
#include 'dcprint.ch'
#include "dccursor.ch"
#include "dcpick.ch"

#include "bdmslib.ch"
#include "bdcolors.ch"
#include "bdbmp.ch"

#include 'amcolors.ch'

#pragma library("asxml10.lib")
#pragma Library("dcxml.lib")
#pragma Library( "asinet10.lib" )
#pragma library( "ascom10.lib" )

#include "Asinet.ch"

#DEFINE CRLF Chr(13)+Chr(10)

* -------------

Function UpClean() // FUNCTION TO CLEAN UP FPT FILES FROM NEWUPS THAT ARE BLOATED

   LOCAL GetList[0], GetOptions
   LOCAL i, nAns
   LOCAL lStatus
   LOCAL aFiles := {'NEWUPS.DBO','NEWUPS.FPO'}

   nAns := Alert_Win(TEXT "Trim Bloated FIle?" ;
                     BUTTONS {"Yes", "No"} ;
                     TITLE "Data Confirmation" )

   IF nAns == 1

     FOR i := 1 TO Len(aFiles)
        IF FExists(aFiles[i])
          IF FErase(aFiles[i]) == -1
            ALERTWIN TEXT "Cannot Erase " + aFiles[i] + ";Please Call Support" TITLE "Data Alert" TIMEOUT 8
          ENDIF
        ENDIF
     NEXT

     IF UseDb( {'NEWUPS'}, 1, .F.)
       NEWUPS->(DbGoTop())
     ELSE
       ALERTWIN TEXT "File In Use - Not Available" TITLE "Data Alert" TIMEOUT 8
       return nil
     ENDIF

    /// COPY FILE TO XNEWUPS THAT WILL NOT INCLUDE BLOATED RECORDS
     COPY TO (GetFlg('ACCTPATH') + 'XNEWUPS')
    CLOSE ALL

     IF FRename(GetFlg('ACCTPATH') + 'NEWUPS.DBF', GetFlg('ACCTPATH') + 'NEWUPS.DBO') == -1
       ALERTWIN TEXT "Cannot Rename NEWUPS.DBF To OLD" + Str(FError()) TITLE "Data Alert" TIMEOUT 5
       return nil
    ENDIF

     IF FRename(GetFlg('ACCTPATH') + 'NEWUPS.FPT', GetFlg('ACCTPATH') + 'NEWUPS.FPO') == -1
       ALERTWIN TEXT "Cannot Rename NEWUPS.FPT To OLD" + Str(FError()) TITLE "Data Alert" TIMEOUT 5
       RETURN NIL
    ENDIF

     IF FRename(GetFlg('ACCTPATH') + 'XNEWUPS.DBF', GetFlg('ACCTPATH') + 'NEWUPS.DBF')==0
       IF Frename(GetFlg('ACCTPATH') + 'XNEWUPS.FPT', GetFlg('ACCTPATH') + 'NEWUPS.FPT')==0
         ALERTWIN TEXT "Clean Up Successful " + Str(FError()) TITLE "Data Alert" TIMEOUT 5
       ELSE
         ALERTWIN TEXT "Clean Up Could Not Be Completed " + Str(FError()) + ";Please Call Support" TITLE "Data Alert" TIMEOUT 5
      ENDIF
    ENDIF
  ENDIF

return nil

* -------------

/// -----------------This function backs into a desired out the door amount
static function _get2desiredups(oRecord,oTabpage4,lOK2Get)
	LOCAL GETLIST[0],GETOPTIONS,nDelta:=0, nAddAmt:=1
	DEFAULT lOK2Get:=.f.
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE SAYFONT '10.Courier New' GETFONT '10.Courier New'
	IF lOK2Get
	 // get the stored desired out the door amount
	 @ 5,0 DCSAY ' Enter the Out The Door Cash Amount To Be Paid By This Customer'
	 @ 6,0 DCSAY 'Note: Any Cash Deficiency Will Be Added To The Capitalized Cost'
	 @ 8,0 DCSAY '                                              Out The Door Cash' GET oRecord:downpay PICTURE '$99999.99'
	 DCREAD GUI MODAL ENTEREXIT ADDBUTTONS FIT OPTIONS GETOPTIONS TITLE 'Desired Out the Door'   OWNER oTabpage4
	ENDIF

	IF oRecord:lterm < 1               /// dont waste time if no term or residual percent
		RETURN oRecord
	ENDIF
	IF oRecord:resper < .01
		return oRecord
	ENDIF

	IF oRecord:quoted < .01
		return oRecord
	ENDIF
	 /// FIRST RUN THE LEASE
	 _makecapcostups(oRecord)

	/// NOW CHECK AGAINST DESIRED OTD PAYMENT
	/// if current inceptions are LESS than the desired amount loop the following by adjusting cashdown by $1 until they are = to pickuppay

	//start with 10 but bump higher when starting at zero for speed
	nAddAmt:=10
	if oRecord:extranum2=0
	 	nAddAmt:=100
	endif


	IF oRecord:incepts < oRecord:downpay
		DO WHILE oRecord:incepts <= oRecord:downpay
	  	   oRecord:extranum2+=nAddAmt
		  _makecapcostups(oRecord)

	   ENDDO
		nAddAmt:=10
			  // if inceptions end up over the pickuppay start adjusting down by a penny

	  IF oRecord:incepts > oRecord:downpay

		  DO WHILE oRecord:incepts >= oRecord:downpay
	  	   oRecord:extranum2-= 1
			_makecapcostups(oRecord)

	     ENDDO
	  ENDIF

	  IF oRecord:incepts < oRecord:downpay

		  DO WHILE oRecord:incepts >= oRecord:downpay
	  	   oRecord:extranum2+= 1
			_makecapcostups(oRecord)

	     ENDDO
	  ENDIF


	  IF oRecord:incepts > oRecord:downpay

		  DO WHILE oRecord:incepts >= oRecord:downpay
	  	   oRecord:extranum2-= .01
			_makecapcostups(oRecord)

	     ENDDO
	  ENDIF

	  IF oRecord:incepts < oRecord:downpay

		  DO WHILE oRecord:incepts >= oRecord:downpay
	  	   oRecord:extranum2+= .01
			_makecapcostups(oRecord)

	     ENDDO
	  ENDIF



	  nDelta:=oRecord:downpay-oRecord:incepts


	  IF nDelta > 0.000
			 oRecord:extranum2+=nDelta
			_makecapcostups(oRecord)

	  ENDIF


	  IF nDelta < 0.000
			 oRecord:extranum2-=nDelta
			 _makecapcostups(oRecord)

	  ENDIF


  ENDIF

  /// if current inceptions are MORE than the desired amount loop the following by adjusting cashdown by -$1 until they are = to pickuppay

  IF oRecord:incepts > oRecord:downpay

   	DO WHILE oRecord:incepts >= oRecord:downpay

	  	  oRecord:extranum2-=nAddAmt
		  _makecapcostups(oRecord)

	   ENDDO
		nAddAmt:=10

 // if inceptions end up less than the pickuppay start adjusting up by a penny

	  IF oRecord:incepts < oRecord:downpay

		  DO WHILE oRecord:incepts <= oRecord:downpay
	  	   oRecord:extranum2+= 1
		  _makecapcostups(oRecord)

	     ENDDO
	  ENDIF

		IF oRecord:incepts > oRecord:downpay

		  DO WHILE oRecord:incepts <= oRecord:downpay
	  	   oRecord:extranum2-= 1
		  _makecapcostups(oRecord)

	     ENDDO
	  ENDIF

	  IF oRecord:incepts < oRecord:downpay

		  DO WHILE oRecord:incepts <= oRecord:downpay
	  	   oRecord:extranum2+= .01
		  _makecapcostups(oRecord)

	     ENDDO
	  ENDIF

		IF oRecord:incepts > oRecord:downpay

		  DO WHILE oRecord:incepts <= oRecord:downpay
	  	   oRecord:extranum2-= .01
		  _makecapcostups(oRecord)

	     ENDDO
	  ENDIF



	  nDelta:=oRecord:downpay-oRecord:incepts

	  IF nDelta < 0.000
			 oRecord:extranum2-=nDelta
			 _makecapcostups(oRecord)

	  ENDIF

	  IF nDelta > 0.000
			 oRecord:extranum2+=nDelta
			 _makecapcostups(oRecord)

	  ENDIF

  ENDIF

  oRecord:capcostred:=(oRecord:extranum2+oRecord:lrebate+(oRecord:tradequote-oRecord:lien))

  IF oRecord:lBuyrate > .001
   _makeleasereserve(oRecord)
  ENDIF
  return oRecord



static function _makecapcostups(oRecord)        //// oRecord:extranum is being used as gross cap cost
	 // get gross capcost

	 IF oRecord:quoted < .01
		RETURN NIL
	 ENDIF

	 oRecord:extranum := oRecord:QUOTED+oRecord:ACQUISFEE+oRecord:servconam+oRecord:aftsaleam1+oRecord:aftsaleam2+;
		oRecord:aftsaleam3+oRecord:aftsaleam4+oRecord:aftsaleam5+oRecord:aftsaleam6

    oRecord:capcost := (oRecord:extranum-(oRecord:extranum2+(oRecord:tradequote-oRecord:lien)+oRecord:lrebate))

	 ///  TO AVOID CRASHING
	 IF valtype(oRecord:addtopymnt) <> 'N'
		oRecord:addtopymnt:=0
	 ENDIF


	/// MAKE LEASE PAYMENT
	_makeleasepayups(@oRecord)


return nil

static function _makeleasepayups(oRecord)                 /// note the oAcedeal:capccr usage is for the tax on capcost reduction - not the first payment
	local nTaxablepays:=0, nTaxableper:=1,nOrigTaxable:=0
   local nEms:=0                        /// extra miles charge   interest rate
	local nNYtax:=1.00, nlPayPart1:=0,nlPayPart2:=0,nLocaltax:=0,nTaxccr:=0,nTaxonpay:=0
	local r:=(oRecord:leaserate1/1200)

	/// start at zero
	oRecord:lPayment:=nlPaypart1:=nlPaypart2:=oRecord:leasetax:=nLocalTax:=oRecord:extranum3:=0                     // extranum3 is property tax
	/// get extra mileage
	nEms:=(oRecord:lTerm/12)*(oRecord:xcessmile*oRecord:Costxcess)
	/// make initial base payment
	IF !oRecord:aprlease
	 nlPayPart1:=((oRecord:Capcost-(oRecord:Residual-nEms)) / oRecord:lTerm )
	 nlPayPart2:=((oRecord:Capcost+(oRecord:Residual-nEms)) * (oRecord:moneyfact1))
	 oRecord:lpayment := nlPaypart1+nlPaypart2
	else
	 oRecord:lPayment:=(oRecord:capcost-(oRecord:Residual-nEms)/(1+r)^oRecord:lTerm)/((1-(1/(1+r)^oRecord:lterm))/r) // calc pay
	 oRecord:lPaypart1:=((oRecord:Capcost-(oRecord:Residual-nEms)) / oRecord:lTerm )
	 oRecord:lPaypart2:=oRecord:lPayment-oRecord:lpaypart1
	ENDIF

	/// calc monthly tax and add on to capcost
	/// IF THIS IS A NY DEAL YOU MUST FACTOR THE TAXES
	IF UPPER(oRecord:STATE)='NY'
	  /// calc monthly lease tax
	  oRecord:leasetax:=oRecord:lPayment*(oRecord:taxper/100)
	  nTaxonPay:=oRecord:leasetax*oRecord:lTerm
	  nTaxCCR:=(oRecord:extranum2+oRecord:LREBATE+oRecord:procfee)*oRecord:taxper/100


	  /// ONLY FACTOR THE TAX FOR NY IF THE TAX IS BEING CAPPED
	  IF oRecord:captax='Y'
	   nNYtax:=1/(1-(oRecord:taxper/100))
	   nTaxonPay:=nTaxonPay*nNytax
		oRecord:capcost+=nTaxonPay

	 /// calc upfront taxes due on cashdown and rebate
		oRecord:capcost+=nTaxCCR
	  endif
	  /// recalc
	 IF !oRecord:aprlease
	  nlPayPart1:=((oRecord:Capcost-(oRecord:Residual-nEms)) / oRecord:lTerm )
	  nlPayPart2:=((oRecord:Capcost+(oRecord:Residual-nEms)) * (oRecord:moneyfact1))
	  oRecord:lpayment := nlPaypart1+nlPaypart2
	  else
	  oRecord:lPayment:=(oRecord:capcost-(oRecord:Residual-nEms)/(1+r)^oRecord:lTerm)/((1-(1/(1+r)^oRecord:lterm))/r) // calc pay
	  oRecord:lPaypart1:=((oRecord:Capcost-(oRecord:Residual-nEms)) / oRecord:lTerm )
	  oRecord:lPaypart2:=oRecord:lPayment-oRecord:lpaypart1
	 ENDIF

    if oRecord:capfees = 'Y'
		 oRecord:capcost+=(oRecord:regfee+oRecord:procfee+oRecord:wastefee+oRecord:inspection)
		 /// recalc
		 IF !oRecord:aprlease
	     nlPayPart1:=((oRecord:Capcost-(oRecord:Residual-nEms)) / oRecord:lTerm )
	     nlPayPart2:=((oRecord:Capcost+(oRecord:Residual-nEms)) * (oRecord:moneyfact1))
	     oRecord:lpayment := nlPaypart1+nlPaypart2
		  else
	     oRecord:lPayment:=(oRecord:capcost-(oRecord:Residual-nEms)/(1+r)^oRecord:lTerm)/((1-(1/(1+r)^oRecord:lterm))/r) // calc pay
	     oRecord:lPaypart1:=((oRecord:Capcost-(oRecord:Residual-nEms)) / oRecord:lTerm )
	     oRecord:lPaypart2:=oRecord:lPayment-oRecord:lpaypart1
	    ENDIF
    endif
	 /// establish final pay as lpayment2
	 oRecord:lpayment2:=oRecord:lPayment+oRecord:addtopymnt

  ENDIF
  /// if not a ny lease
  IF UPPER(oRecord:STATE) <> 'NY'
			/// calc upfront taxes on ccr and tot fees

			/// calc upfront taxes due on cashdown and rebate
	      nTaxCCR:=(oRecord:extranum2+oRecord:LREBATE+oRecord:procfee)*oRecord:taxper/100


			IF oRecord:Captax='Y'
				oRecord:capcost+=nTaxCCR
			ENDIF


			if oRecord:capfees = 'Y'
			  oRecord:capcost+=(oRecord:regfee+oRecord:procfee+oRecord:wastefee+oRecord:inspection)
			endif


			//recalc
			IF !oRecord:aprlease
			 nlPayPart1:=((oRecord:Capcost-(oRecord:Residual-nEms)) / oRecord:lTerm )
	       nlPayPart2:=((oRecord:Capcost+(oRecord:Residual-nEms)) * (oRecord:moneyfact1))
	       oRecord:lpayment := nlPaypart1+nlPaypart2
			else
	       oRecord:lPayment:=(oRecord:capcost-(oRecord:Residual-nEms)/(1+r)^oRecord:lTerm)/((1-(1/(1+r)^oRecord:lterm))/r) // calc pay
	       oRecord:lPaypart1:=((oRecord:Capcost-(oRecord:Residual-nEms)) / oRecord:lTerm )
	       oRecord:lPaypart2:=oRecord:lPayment-oRecord:lpaypart1
	      ENDIF


			//oAceDeal:totcapcost := oAceDeal:capcost+oAcedeal:leasetax
			oRecord:leasetax:=(oRecord:lpayment)*(oRecord:taxper/100)
			  /// GIVE SALES TAX CREDIT ON LIENPAYOFF

			IF upper(oRecord:state)='CT'
				/// zero out prop tax before recalcing
				oRecord:extranum3:=0


				/// now calc prop tax if extrachar3='Y' 									                                   //extrachar3 and extranum3 are for property taxes
				IF alltrim(oRecord:extrachar3) = 'Y'     /// calc prop tax
					IF oRecord:tradequote-oRecord:lien  < 0

					 oRecord:extranum3 := (oRecord:Quoted+(oRecord:tradequote-oRecord:lien)) *.001225 /// tax on proptax

					else
					 oRecord:extranum3 := (oRecord:Quoted) *.001225 /// tax on proptax
					ENDIF
					IF oRecord:extranum3 < .01
						oRecord:extranum3:=0
					ENDIF
					oRecord:leasetax:=(oRecord:lpayment+oRecord:extranum3)*(oRecord:taxper/100)

				endif
				/// calc if owe on trade exists for conn per state tax rules
				/// calc total sales tax due over the term
				nOrigTaxable:=(oRecord:lPayment)*oRecord:lterm
				//deduct from total sales tax due the amount owed on the trad

				/// calc percent to use as a factor

				IF oRecord:lien-oRecord:tradequote > 0
				 nTaxablePays:=nOrigTaxable-(oRecord:lien-oRecord:tradequote)
				ELSE
				 nTaxablePays:=nOrigTaxable
				ENDIF
				/// calc percent to use as a factor


				nTaxableper:=nTaxablePays/nOrigTaxable
				/// factor the original monthly tax payment
				oRecord:leasetax:=oRecord:leasetax*(nTaxableper)

			ENDIF

	 //wtf oRecord:lpayment,oRecord:leasetax
	 oRecord:lpayment2:=oRecord:lPayment+oRecord:leasetax+oRecord:extranum3+oRecord:addtopymnt

	 //wtf oRecord:lpayment2,oRecord:lPayment,oRecord:leasetax

	ENDIF   /// not ny


	oRecord:incepts := oRecord:EXTRANUM2+oRecord:lPayment2
	//wtf oRecord:incepts
	IF oRecord:capfees='N'
		oRecord:incepts:=oRecord:incepts+(oRecord:regfee+oRecord:procfee+oRecord:wastefee+oRecord:inspection)
	ENDIF

	IF oRecord:captax ='N'
		oRecord:incepts:=oRecord:incepts+(nTaxonpay+nTaxccr)
	ENDIF

	oRecord:incepts:=oRecord:incepts

	/// we may want to make this optionnal to pad the dueatsign with first payment since it is already in the cash down
	IF iniget('FNI','NOFIRSTPAY')='Y'
		oRecord:DUEATSIGN:=oRecord:incepts
	  ELSE
	   oRecord:DUEATSIGN:=oRecord:lpayment2+oRecord:incepts

	ENDIF


	return oRecord:Capcost



static function _gettoleasepayups(oRecord,oTabpage4)
  local GETLIST[0],GETOPTIONS,oAcedealNew ,xStatus
  local nAmtreduce:=0, nNewPay:=0,nNewCapCost:=0,oAdjMF,oToolbar5,oToolbar6 , lOK2Get:=.f.
  local nApptosale:=0,nApptodown:=0 ,oAdjPrice,oToolbar,oToolbar2 ,oAdjdown,oToolbar3,oToolbar4
  DCGETOPTIONS SAYWIDTH 0 SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE

  oAceDealNew:=NEWUPS->(DC_DBRECORD():NEW())
  DC_DBRECORDCOPY(oRecord,oAcedealNew)
  oAceDealNew:aprlease:= ACELEASE->aprlease
  oAceDealNew:addtopymnt:=ACELEASE->addtopymnt



   @ 1,1 dcsay 'Monthly Payment' get oRecord:lPayment2       picture '$999999.99'    editprotect{|| .t.}     //PARENT oAdjPrice
	@ 2,1 dcsay 'Current CapCost' get oRecord:capcost        picture '$999999.99'    editprotect{|| .t.}     //PARENT oAdjPrice
	//@ 2,50 DCSAY 'Current Vehicle Gross'  GET oAceDeal:vehgross  PICTURE '$999999.99' editprotect{|| .t.}  GETCOLOR{|| IIF( oAceDeal:vehgross > 0,{1,BD_KEYLIME} ,{1,BD_SALMON} )}

	@ 3,0 DCGROUP oAdjPrice CAPTION 'Adjust Price By Button'  SIZE 120,6.5

		@ 3,85 DCSAY '    Selling Price' get oRecord:quoted       picture '$999999.99' editprotect{|| .t.}     PARENT oAdjPrice

	   @ 1,1 DCTOOLBAR oToolbar SIZE 100,2 BUTTONSIZE 10,2  ;
		PARENT oAdjPrice

	  DCADDBUTTON CAPTION 'Add .01' ;
		  PARENT oToolbar ;
		  ACTION {|| oRecord:quoted+=.01 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}
	  DCADDBUTTON CAPTION 'Add .10' ;
		  PARENT oToolbar ;
		  ACTION {|| oRecord:quoted+=.10 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}
	  DCADDBUTTON CAPTION 'Add .50' ;
		  PARENT oToolbar ;
		  ACTION {|| oRecord:quoted+=.50 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	  DCADDBUTTON CAPTION 'Add 1.00' ;
		  PARENT oToolbar ;
		  ACTION {|| oRecord:quoted+=1.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Add 5.00' ;
		  PARENT oToolbar ;
		  ACTION {|| oRecord:quoted+=5.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Add 10.00' ;
		  PARENT oToolbar ;
		  ACTION {|| oRecord:quoted+=10.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Add 50.00' ;
		  PARENT oToolbar ;
		  ACTION {|| oRecord:quoted+=50.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Add 100' ;
		  PARENT oToolbar ;
		  ACTION {|| oRecord:quoted+=100 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	  @ 4,1 DCTOOLBAR oToolbar2 SIZE 100,2 BUTTONSIZE 10,2  ;
		PARENT oAdjPrice

	  DCADDBUTTON CAPTION 'Deduct .01' ;
		  PARENT oToolbar2 ;
		  ACTION {|| oRecord:quoted-=.01 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}
	  DCADDBUTTON CAPTION 'Deduct .10' ;
		  PARENT oToolbar2 ;
		  ACTION {|| oRecord:quoted-=.10 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}
	  DCADDBUTTON CAPTION 'Deduct .50' ;
		  PARENT oToolbar2 ;
		  ACTION {|| oRecord:quoted-=.50 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	  DCADDBUTTON CAPTION 'Deduct 1.00' ;
		  PARENT oToolbar2 ;
		  ACTION {|| oRecord:quoted-=1.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Deduct 5.00' ;
		  PARENT oToolbar2 ;
		  ACTION {|| oRecord:quoted-=5.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Deduct 10.00' ;
		  PARENT oToolbar2 ;
		  ACTION {|| oRecord:quoted-=10.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}
	 DCADDBUTTON CAPTION 'Deduct 50.00' ;
		  PARENT oToolbar2 ;
		  ACTION {|| oRecord:quoted-=50.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Deduct 100' ;
		  PARENT oToolbar2 ;
		  ACTION {|| oRecord:quoted-=100 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}


	@ 10,0 DCGROUP oAdjDown CAPTION ' Adjust Out The Door Cash By Button'  SIZE 120,7
	@ 1,1 DCTOOLBAR oToolbar3 SIZE 100,2 BUTTONSIZE 10,2  ;
		PARENT oAdjDown
	@ 3,85 DCSAY 'Out The Door Cash' get oRecord:downpay   picture '$999999.99' editprotect{|| .t.}     PARENT oAdjDown

	  DCADDBUTTON CAPTION 'Add .01' ;
		  PARENT oToolbar3 ;
		  ACTION {|| oRecord:downpay+=.01 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	  DCADDBUTTON CAPTION 'Add .10' ;
		  PARENT oToolbar3 ;
		  ACTION {|| oRecord:downpay+=.10 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}
	  DCADDBUTTON CAPTION 'Add .50' ;
		  PARENT oToolbar3 ;
		  ACTION {|| oRecord:downpay+=.50 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	  DCADDBUTTON CAPTION 'Add 1.00' ;
		  PARENT oToolbar3 ;
		  ACTION {|| oRecord:downpay+=1.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Add 5.00' ;
		  PARENT oToolbar3 ;
		  ACTION {|| oRecord:downpay+=5.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Add 10.00' ;
		  PARENT oToolbar3 ;
		  ACTION {|| oRecord:downpay+=10.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}
	 DCADDBUTTON CAPTION 'Add 50.00' ;
		  PARENT oToolbar3 ;
		  ACTION {|| oRecord:downpay+=50.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Add 100' ;
		  PARENT oToolbar3 ;
		  ACTION {|| oRecord:downpay+=100 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	@ 4,1 DCTOOLBAR oToolbar4 SIZE 100,2 BUTTONSIZE 10,2  ;
		PARENT oAdjDown

	  DCADDBUTTON CAPTION 'Deduct .01' ;
		  PARENT oToolbar4 ;
		  ACTION {|| oRecord:downpay-=.01 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	  DCADDBUTTON CAPTION 'Deduct .10' ;
		  PARENT oToolbar4 ;
		  ACTION {|| oRecord:downpay-=.10 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}
	  DCADDBUTTON CAPTION 'Deduct .50' ;
		  PARENT oToolbar4 ;
		  ACTION {|| oRecord:downpay-=.50 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	  DCADDBUTTON CAPTION 'Deduct 1.00' ;
		  PARENT oToolbar4 ;
		  ACTION {|| oRecord:downpay-=1.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Deduct 5.00' ;
		  PARENT oToolbar4 ;
		  ACTION {|| oRecord:downpay-=5.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Deduct 10.00' ;
		  PARENT oToolbar4 ;
		  ACTION {|| oRecord:downpay-=10.00 ,;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}
	 DCADDBUTTON CAPTION 'Deduct 50.00' ;
		  PARENT oToolbar4 ;
		  ACTION {|| oRecord:downpay-=50.00 ,;
				 _get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Deduct 100' ;
		  PARENT oToolbar4 ;
		  ACTION {|| oRecord:downpay-=100 ,;
				 _get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}


	@ 17,0 DCGROUP oAdjMF CAPTION ' Adjust Interest Rate / Money Factor By Button'  SIZE 120,7.5

	@ 1,1 DCTOOLBAR oToolbar5 SIZE 80,2 BUTTONSIZE 10,2  ;
		PARENT oAdjMF

	@ 3,85 DCSAY '    Interest Rate' get oRecord:leaserate1   picture '999.99' editprotect{|| .t.}     PARENT oAdjMF
	@ 6,85 DCSAY '     Money Factor' get oRecord:moneyfact1   picture '9.999999' editprotect{|| .t.}     PARENT oAdjMF


	  DCADDBUTTON CAPTION 'Add .01' ;
		  PARENT oToolbar5 ;
		  ACTION {|| oRecord:leaserate1+=.01 ,;
					makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	  DCADDBUTTON CAPTION 'Add .05' ;
		  PARENT oToolbar5 ;
		  ACTION {|| oRecord:leaserate1+=.05 ,;
					 makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}


	  DCADDBUTTON CAPTION 'Add .10' ;
		  PARENT oToolbar5 ;
		  ACTION {|| oRecord:leaserate1+=.10 ,;
					makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	  DCADDBUTTON CAPTION 'Add .25' ;
		  PARENT oToolbar5 ;
		  ACTION {|| oRecord:leaserate1+=.25 ,;
					makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Add .50' ;
		  PARENT oToolbar5 ;
		  ACTION {|| oRecord:leaserate1+=.50 ,;
					makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Add .75' ;
		  PARENT oToolbar5 ;
		  ACTION {|| oRecord:leaserate1+=.75 ,;
					makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Add 1.00' ;
		  PARENT oToolbar5 ;
		  ACTION {|| oRecord:leaserate1+=1.00 ,;
					makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Add 1.25' ;
		  PARENT oToolbar5 ;
		  ACTION {|| oRecord:leaserate1+=1.25 ,;
					makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	@ 4,1 DCTOOLBAR oToolbar6 SIZE 100,2 BUTTONSIZE 10,2  ;
		PARENT oAdjMF

	  DCADDBUTTON CAPTION 'Deduct .01' ;
		  PARENT oToolbar6 ;
		  ACTION {|| oRecord:leaserate1-=.01 ,;
					makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	  DCADDBUTTON CAPTION 'Deduct .05' ;
		  PARENT oToolbar6 ;
		  ACTION {|| oRecord:leaserate1-=.05 ,;
					makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	  DCADDBUTTON CAPTION 'Deduct .10' ;
		  PARENT oToolbar6 ;
		  ACTION {|| oRecord:leaserate1-=.10 ,;
					makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	  DCADDBUTTON CAPTION 'Deduct .25' ;
		  PARENT oToolbar6 ;
		  ACTION {|| oRecord:leaserate1-=.25 ,;
					 makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Deduct .50' ;
		  PARENT oToolbar6 ;
		  ACTION {|| oRecord:leaserate1-=.50 ,;
					makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Deduct .75' ;
		  PARENT oToolbar6 ;
		  ACTION {|| oRecord:leaserate1-=.75 ,;
					makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Deduct 1.00' ;
		  PARENT oToolbar6 ;
		  ACTION {|| oRecord:leaserate1-=1.0 ,;
					 makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}

	 DCADDBUTTON CAPTION 'Deduct 1.25' ;
		  PARENT oToolbar6 ;
		  ACTION {|| oRecord:leaserate1-=1.25 ,;
					 makemoneyfact(@oRecord),;
					_get2desiredups(oRecord,oTabpage4,lOK2Get),;
					DC_GETREFRESH(GETLIST)}


  	@ 30,1 dcsay '                          Enter Desired Payment' get nNewPay picture '$99999.99'  VALID{ | | _makenewleaseups(nNewpay,@nNewCapCost,@nAmtreduce,oAceDealNew,oRecord),;
	 	dc_getrefresh(getlist),.t.} NOTABSTOP
  	@ 31,1 dcsay 'Amount of CAPCOST Change Needed To Get Payment.' get nAmtreduce picture '$99999.99' ;//editprotect{|| .t.};
	 GETCOLOR{|| IIF( nAmtReduce > 0,{1,BD_KEYLIME} ,{1,BD_SALMON} )  }

	dcread gui to xstatus fit addbuttons options getoptions title 'Get to Payment' OWNER oTabpage4
	if !xstatus
		return nil
	endif
return nil

static function _makenewleaseups(nNewpay,nNewCapcost,nAmtreduce,oAcedealNew,oRecord)

	IF nNewpay < oAceDealNew:lpayment2
			DO WHILE oAceDealNew:lpayment2 > nNewpay
				oAcedealNew:quoted += -100
				_get2desiredUps(@oAcedealNew,'',.F.)
			enddo
			DO WHILE oAceDealNew:lpayment2 < nNewpay
				oAcedealNew:quoted  += 10
				_get2desiredUps(@oAcedealNew,'',.F.)
			enddo

	 else
		  DO WHILE nNewpay > oAceDealNew:lpayment2
				oAcedealNew:quoted  += 100
				_get2desiredUps(@oAcedealNew,'',.F.)
			enddo
			DO WHILE nNewpay > oAceDealNew:lpayment2
				oAcedealNew:quoted += -10
				_get2desiredUps(@oAcedealNew,'',.F.)
			enddo
	ENDIF

	nAmtReduce:=oAcedealNew:capcost-oRecord:capcost
	RETURN nAmtReduce





FUNCTION PURGEAPPRSL()
	LOCAL GETLIST:={},GETOPTIONS, dStartDate:=DATE()-90,oRecord, xStatus ,oDialog
	DCGETOPTIONS SAYWIDTH 0 SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
	IF UseDb({'UCAPPRSL','HUCAPPRSL'},1,.F.)
		UCAPPRSL->(DBGOTOP())
	ELSE
		DC_WINALERT('Cannot open all Appraisal files to perform purge')
		CLOSE ALL
		RETURN NIL
	ENDIF

	@ 10,0 DCSAY 'This program will copy appraisal files to history that fall before the '
	@ 11,0 DCSAY 'below entered date.'

	@ 12,0 DCSAY 'Copy appraisals to history that are dated before' get dStartdate
	DCREAD GUI TO xStatus BUTTONS 2 ENTEREXIT OPTIONS GETOPTIONS FIT TITLE 'Archive Appraisal Records'

	SELECT UCAPPRSL
	UCAPPRSL->(DBGOTOP())
	oDialog:=DC_WAITON('Now copying records.')
	DO WHILE !UCAPPRSL->(EOF())

		IF UCAPPRSL->APPRDATE < dStartdate
			oRecord:=UCAPPRSL->(DC_DBRECORD():NEW())
			UCAPPRSL->(DC_DBSCATTER(oRecord))
			SELECT HUCAPPRSL
			HUCAPPRSL->(DC_DBGATHER(oRecord,.t.))
			SELECT UCAPPRSL
			UCAPPRSL->(DBDELETE())
		ENDIF
		UCAPPRSL->(DBSKIP(1))
	ENDDO

	SLEEP(100)
	SELECT UCAPPRSL
	UCAPPRSL->(DBPACK())

	DC_IMPL(oDialog)
	CLOSE ALL
	RETURN NIL

static function _SENDUPTOVIN(nRec)
	LOCAL cBody:='',cSource:='' ,ceAddress,cVendor,GETLIST:={},GETOPTIONS,xStatus
	local cExdate:=dtoS(date())
   local cExtime:=TIME()
	LOCAL cMake:=SPACE(15),cModel:=SPACE(15),cYear:=SPACE(4),cVin:='',cTrim:='',cSM:='',cCity,cState,cStatus,cInterest:='"buy"'
	LOCAL cTrMake:='',cTrModel:='',cTrYear:='',cTrVin:='',cTrTrim:='',cTrMileage:='',cTrAppraisal:=''
	DCGETOPTIONS SAYWIDTH 0 SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
	SELECT NEWUPS
	NEWUPS->(DBGOTO(nRec))
	IF !EMPTY(NEWUPS->NEWUSED)
		IF !EMPTY(NEWUPS->STKQUOTED)
			_GETVEHINFO(NEWUPS->NEWUSED,NEWUPS->STKQUOTED,@cYear,@cMake,@cModel,@cVin,@cTrim)
		ENDIF
	ENDIF

	IF empty(cYear)
		@ 1,0 DCSAY 'Vin Solutions REQUIRES Some Vehicle Interest Data'
		@ 2,0 DCSAY 'Since You have not entered a Stock Number Please enter below Info.'
		@ 3,0 DCSAY 'If you want to go back and fill in this data with a Stock Number press ESC '
		@ 4,0 DCSAY 'OR ELSE FILL IT IN HERE'
		@ 6,0 DCSAY 'Year  '   get cYear  valid{|| IIF( empty(cYear),.f. ,.t. ) }
		@ 7,0 DCSAY 'Make  '   get cMake  valid{|| IIF( empty(cMake),.f. ,.t. ) }
		@ 8,0 DCSAY 'Model '   get cModel valid{|| IIF( empty(cModel),.f. ,.t. ) }
		DCREAD GUI MODAL TO xStatus ENTEREXIT OPTIONS GETOPTIONS TITLE 'Vehicle Info for Vin' EVAL{|o| SETAPPWINDOW(o)}
		IF !xStatus
		 RETURN NIL
	   ENDIF

	   IF empty(cYear)
		 DC_WINALERT('Record Not sent. Not enough Info')
		 RETURN NIL
	   ENDIF


	ENDIF


	getzip(NEWUPS->ZIPCODE,@cCity,@cState)

	SELECT SM
	SM->(DBSEEK(NEWUPS->SALESMAN))
	IF FOUND()
		cSM:=ALLTRIM(SM->INETNAME)
	ENDIF
	IF !empty(NEWUPS->ADSOURCE)
		SELECT ADSOURCE
		DBSEEK(NEWUPS->ADSOURCE)
		IF FOUND()
			cSource:=alltrim(ADSOURCE->DESC)
			cSource:='"'+cSource+'"'
		else
			cSource:='Floor_Traffic'
		ENDIF
	ENDIF
	IF usedb({'VINLEADS'})
		VINLEADS->(DBGOTOP())
	ELSE
		return nil
	ENDIF

	IF NEWUPS->NEWUSED='N'
		cStatus:='"new"'
	else
		cStatus:='"used"'
	ENDIF
	cTrMake:=ALLTRIM(NEWUPS->TRADEMAKE)
	cTrModel:=ALLTRIM(NEWUPS->TRADEDESC)
	cTrYear:=NEWUPS->TRADEYR
	cTrVin:=NEWUPS->TRADEVIN
	cTrMileage:=ALLTRIM(STR(NEWUPS->TRADEMILES))
	cTrAppraisal:=alltrim(str(NEWUPS->TRADEQUOTE))
	ceAddress:=alltrim(VINLEADS->EMAIL)
	cVendor:=alltrim(VINLEADS->COMPANY)

	VINLEADS->(DBCLOSEAREA())

	cExDate:=substr(cExdate,1,4)+'-'+substr(cExdate,5,2)+'-'+substr(cExdate,7,2)

	SELECT NEWUPS

	cBody=cBody+'<?XML VERSION=“1.0” encoding="UTF-8"?>'+CRLF
	cBody=cBody+'<?adf version="1.0" ?>'+CRLF
   cBody=cBody+'<adf>'+CRLF
   cBody=cBody+' <prospect>  '+CRLF
   cBody=cBody+'  <id sequence="1" source='+cSource+'</id>'+CRLF
   cBody=cBody+'  <requestdate>'+cExdate+'T'+cExtime+'-05:00'+'</requestdate> '+CRLF
   cBody=cBody+'  <vehicle interest='+cInterest+' status='+cStatus+'> '+CRLF
   cBody=cBody+'   <id sequence="1" source='+cSource+'>'+NEWUPS->UPID+'</id>   '+CRLF
   cBody=cBody+'   <year>'+cYear+'</year>'+CRLF
   cBody=cBody+'   <make>'+ALLTRIM(cMake)+'</make>'+CRLF
   cBody=cBody+'   <model>'+ALLTRIM(cModel)+'</model>'+CRLF
   cBody=cBody+'   <vin>'+cVin+'</vin>'+CRLF
   cBody=cBody+'   <stock>'+ALLTRIM(NEWUPS->STKQUOTED)+'</stock>'+CRLF
   cBody=cBody+'   <trim>'+ALLTRIM(cTrim)+'</trim>'+CRLF
   cBody=cBody+'   <price type="asking" currency="USD"></price>  '+CRLF
   cBody=cBody+'   </vehicle>                                     '+CRLF
	IF !empty(NEWUPS->TRADEVIN)
		cBody=cBody+'  <vehicle interest="trade-in" status=""> '+CRLF
      cBody=cBody+'   <id sequence="1" source='+cSource+'>'+NEWUPS->UPID+'</id>   '+CRLF
   	cBody=cBody+'   <year>'+cTrYear+'</year>'+CRLF
   	cBody=cBody+'   <make>'+cTrMake+'</make>'+CRLF
   	cBody=cBody+'   <model>'+cTrModel+'</model>'+CRLF
   	cBody=cBody+'   <vin>'+cTrVin+'</vin>'+CRLF
   	cBody=cBody+'   <stock></stock>'+CRLF
		cBody=cBody+'   <odometer status="unknown" units="mi">'+cTrMileage+'</odometer>'+CRLF
   	cBody=cBody+'   <trim>'+cTrTrim+'</trim>'+CRLF
   	cBody=cBody+'   <price type="appraisal" currency="USD">'+cTrAppraisal+'</price>' +CRLF
   	cBody=cBody+'   </vehicle>'+CRLF
	ENDIF
   cBody=cBody+'  <customer>                                     '+CRLF
   cBody=cBody+'    <contact>                                    '+CRLF
   cBody=cBody+'    <name part="first">'+ALLTRIM(NEWUPS->FIRSTNAME)+'</name> '+CRLF
   cBody=cBody+'    <name part="last">'+ALLTRIM(NEWUPS->LASTNAME)+'</name>'+CRLF
   cBody=cBody+'    <email>'+ALLTRIM(NEWUPS->EMAIL)+'</email> '+CRLF
   cBody=cBody+'    <phone type="voice" time="day">'+NEWUPS->AREA+NEWUPS->UPID+'</phone>      '+CRLF
	cBody=cBody+'    <address type="home" >'+CRLF
	cBody=cBody+'     <postalcode>'+NEWUPS->ZIPCODE+'</postalcode>'+CRLF
	cBody=cBody+'     <street line="1">'+alltrim(NEWUPS->STREET)+'</street>      '+CRLF
	cBody=cBody+'     <city>'+cCity+'</city>'+CRLF
	cBody=cBody+'     <regioncode>'+cState+'</regioncode>'+CRLF
	cBody=cBody+'    </address>'                                            +CRLF
   cBody=cBody+'   </contact>                                              '+CRLF
   cBody=cBody+'   <comments>This is an AutoMan generated Prospect</comments>'+CRLF
   cBody=cBody+' </customer>'+CRLF
   cBody=cBody+' <vendor>'+CRLF
   cBody=cBody+'  <id>vinlens</id>                                             '+CRLF
   cBody=cBody+'  <vendorname>'+cVendor+' [Attn: '+cSM+']</vendorname>   '+CRLF
   cBody=cBody+'  <contact>                                                         '+CRLF
   cBody=cBody+'    <name/>                                                         '+CRLF
   cBody=cBody+'    <phone type="voice" time="nopreference">8452258468</phone>      '+CRLF
   cBody=cBody+'  </contact>                                                        '+CRLF
   cBody=cBody+' </vendor>                                                          '+CRLF
   cBody=cBody+' <provider>                                                         '+CRLF
   cBody=cBody+'   <name>Automan - </name>                                          '+CRLF
   cBody=cBody+'   <service>Automan</service>                                       '+CRLF
   cBody=cBody+'   <url>http://www.meadowlandsystems.com</url>                           '+CRLF
   cBody=cBody+'   <email>info@meadowlandsystems.com</email>                          '+CRLF
   cBody=cBody+'   <phone>845-225-8468</phone>                                      '+CRLF
   cBody=cBody+' </provider>                                                        '+CRLF
   cBody=cBody+' <leadtype>Floor_Traffic</leadtype>                                 '+CRLF
   cBody=cBody+'</prospect>                                                         '+CRLF
   cBody=cBody+'</adf>                                                              '+CRLF

	_EmailUpToVin(SMTPSERVER(),'bvolz3850@gmail.com',ceAddress,'Lead Upload','bvolz3850@gmail.com',' ',cBody,;
		.t.,' ', '','','', .t. ,'',.t.,.t.,'')


	RETURN NIL

static function getzip(zip,cCity,cState)
select usazip
seek zip
IF !FOUND()
  cCity:='Unknown'
  cState:='NA'
ENDIF
if found()
  cCity:=alltrim(usazip->city)
  cState:=alltrim(usazip->state)
endif
return .t.


 static Function _EmailUpToVin(cSMTPSERVER,cSender,cRecipient,cSubject,cReplyto,cAttachment,cText,;
 	lSendMail, cCC, cBCC,cBCC2,cBCC3, lDisplayMsg,cPassword,lConfirm,lSentMessOk,cCannedMess)
	LOCAL oMail, oSender, oRecipient, oSmtp, lError := .t., i ,oCC, nPort:=25 ,aStruct:={}
	LOCAL lConnectOK   ,cAuthUser:=space(1) , cAuthPass:=space(1) ,lSendOK:=.f., oBCC,oBCC2,oBCC3, oMSG
	local aArray:={cSmtpserver,cSender,cRecipient,cSubject,cReplyto,cAttachment,cText,lSendmail,cCC,cBCC,cBCC3,lDisplaymsg,cPassword,lConfirm,lSentmessOK,cCannedmess}


	DEFAULT cSubject TO PAD('eMail  ',50) , ;
		cReplyTo TO SPACE(50) , ;
		cText TO '' , ;
		cAttachment TO SPACE(50), ;
		lSendMAil TO .t., ;
		cCC TO '', ;
		cBCC TO '', ;
		lDisplayMsg TO .T.,;
		lConfirm to .t. , ;
		lSentMessOK to .t. ,;
		cCannedMess to ''

	   cSubject := strtran(cSubject,":","-")

	/// send as text with html

	//FixText(@cText)

	IF !EMPTY(iniget('SYSTEM','VERIZON'))
			cSender:=iniget('SYSTEM','VERIZON')
	ENDIF



	// cText:='<pre>'+cText+'</pre>'

	IF lSendMail

		oMail := MIMEMessage():new()
		oSender := MailAddress():new( '<'+lower(Alltrim(cSender))+'>' )

		// Assemble the e-mail

		oRecipient := MailAddress():new( '<'+lower(Alltrim(cRecipient))+'>' )
		oMail:addRecipient( oRecipient )

		oMail:setFrom( oSender )
		oMail:setSubject( Alltrim(cSubject) )
		oMail:setMessage( cText )
		oMail:addHeader( "Reply-To", '<'+lower(Alltrim(cReplyTo))+'>' )
		IF !empty(cCC)
			oMail:addHeader( "CC", '<'+lower(Alltrim(cCC))+'>' )
		ENDIF
		oMail:setContentType('text/html')

		IF !Empty(cAttachment)
			oMail:attachFile( cAttachment )
		ENDIF

		//wtf cAttachment
		//oSmtp := smTPClient():new( Alltrim(cSMTPServer) )  moved below 7-11-14

	  IF !FEXISTS(GETFLG('HISTPATH')+'EMAILAUT.DBF')
		  aStruct:={}
		  aAdd(aStruct,{'USERNAME','C',50,0})
		  aAdd(aStruct,{'PASSWORD','C',50,0})
		  aAdd(aStruct,{'SMTPPORT','N',3,0})
		  aAdd(aStruct,{'POPPORT','N',3,0})
		  DBCREATE('EMAILAUT',aStruct,'FOXCDX')
	  ENDIF



		//IF fexists('f:\aman\EMAILAUT.DBF')                        /// this will get password for mail if needed
		IF UseDB('EMAILAUT')
			cAuthuser:=alltrim(EMAILAUT->USERNAME)
			cAuthpass:=alltrim(EMAILAUT->PASSWORD)
			nPort:=EMAILAUT->SMTPPORT
			EMAILAUT->(DBCLOSEAREA())
		ELSE
			DC_WINALERT('No email authorization available.')
			//lconnectok:=oSmtp:connect(cAuthuser,cAuthpass)
			//else
		   //lconnectok:=oSmtp:connect()
		ENDIF


		IF nPort=25 .OR. nPort=0                                                     /// let default be 25 or else send port number
		 oSmtp := smTPClient():new( Alltrim(cSMTPServer) )
		else
		 oSmtp := smTPClient():new( Alltrim(cSMTPServer), nPort )
		endif



		IF !empty(cAuthuser)                                            /// only send with userid and pw if not empty
		 lconnectok:=oSmtp:connect(cAuthuser,cAuthpass)
			else
		 lconnectok:=oSmtp:connect()
		endif


		// IF oSmtp:connect()
		IF lConnectOK
			lsendok:=oSmtp:send( oMail )
			oSmtp:disconnect()

			///wtf lSendok

			IF lSendok
			  lSentMessOk:=.t.
			 _UPDATEMAILLOG(cSender,cRecipient,cSubject,cText,cCannedMess,cSMTPServer,cCC,cBCC)
			ENDIF
			if lConfirm
				if lsendok
					lSentmessOk:=.t.
					dc_winalert('Mail sent OK','eMail results.',XBPMB_OK,XBPMB_INFORMATION)
				  else
					dc_winalert('Sorry Mail Not Sent!!')
				endif
			endif
		ELSE
			if lConfirm
				IF lDisplayMsg
					dc_winalert('Unable to connect to mail server. Please try later.','eMail Results.',XBPMB_OK,XBPMB_WARNING)
				ENDIF
			endif
			lError := .F.
		ENDIF

		SET TIME TO HH:MM:SS
	ENDIF

RETURN lError

static function _UPDATEMAILLOG(cSender,cRecipient,cSubject,cText,cCannedMess,cSMTPServer,cCC,cBCC)
	 select EMAILLOG
	 IF EMAILLOG->(DC_ADDREC())
		REPLACE EMAILLOG->EMAIL WITH cRecipient
		REPLACE EMAILLOG->SENDER WITH cSender
		REPLACE EMAILLOG->SUBJECT WITH cSubject
		REPLACE EMAILLOG->CANNEDMESS WITH cCannedMess
		REPLACE EMAILLOG->SMTPSERVER WITH cSmtpserver
		REPLACE EMAILLOG->MESSAGE WITH ctext
		REPLACE EMAILLOG->DATE WITH DATE()
		REPLACE EMAILLOG->TIME WITH TIME()
		REPLACE EMAILLOG->CCTO WITH cCC
		REPLACE EMAILLOG->BCCTO WITH cBCC
	 ENDIF
	RETURN NIL





 static function _GETVEHINFO(cNewUsed,cStockno,cYear,cMake,cModel,cVin,cTrim)
	IF cNewused='N'
		SELECT NCI
		NCI->(DBSEEK(cStockno))
		IF FOUND()
			cYear:=NCI->YEAR
			cMake:=NCI->MAKE
			cModel:=NCI->MODEL
			cVin:=NCI->VIN
			cTrim:=NCI->TRIM
		ENDIF
	ENDIF
	IF cNewused='U'
		SELECT UCL01
		UCL01->(DBSEEK(cStockno))
		IF FOUND()
			cYear:=UCL01->YEAR
			cMake:=UCL01->MAKE
			cModel:=UCL01->MODEL
			cVin:=UCL01->VIN
			cTrim:=UCL01->TRIM
		ENDIF
	ENDIF
RETURN NIL


function browapptbysm()
	local getlist:={},getoptions,nRow,nCol
	LOCAL xStatus, nRec,oConfig1,oConfig2,oDialog, cSM:=space(3)
	LOCAL oBrowse,apres,aColPres,totoolbar,dQdate:=DATE(), aPush:={}, aChars:={} ,dEnddate:=DATE()+1, aArray:={}
	LOCAL aWaitcolor:={GRA_CLR_BLACK,GRA_CLR_WHITE}
	//SET DELETED ON
	DCGEToptions autoresize saywidth 0 SAYFONT('8.Courier New') getfont('8.Courier New')

	oConfig1 := DC_XbpPushButtonXPConfig():new()
	oConfig1:bitmapOffset := 5
	oConfig1:fgColorMouse := COLOR_BLACK
	oConfig1:bgColorMouse := COLOR_ORANGE
	oConfig1:fgColor := COLOR_BLACK
	oConfig1:bgColor := BD_LINEN
	oConfig1:gradientStep := 10
	oConfig1:gradientReverse := .t.
	oConfig1:radius := 10
	oConfig1:bordercolor := COLOR_RED

	oConfig2 := DC_XbpPushButtonXPConfig():new()
	oConfig2:bitmapOffset := 5
	oConfig2:fgColorMouse := COLOR_BLACK
	oConfig2:bgColorMouse := COLOR_ORANGE
	oConfig2:fgColor := COLOR_BLACK
	oConfig2:bgColor := COLOR_SILVER
	oConfig2:gradientStep := 10
	oConfig2:gradientReverse := .t.
	oConfig2:radius := 10
	oConfig2:bordercolor := COLOR_BLUE



	oDialog:=DC_WAITON('Now Opening Files')
	IF !OPENROFILES()
		DC_IMPL(oDialog)
		RETURN NIL
	ENDIF
	if !OPENSAKHISTFILES()
	   DC_IMPL(oDialog)
	   return nil
   endif
	DC_IMPL(oDialog)

	IF !UseDb({'VTMAS','TEAM'})
	  DC_WINALERT('Cannot Open VTMAS OR TEAM')
	  return nil
   ENDIF




	SERVRESV->(ORDSETFOCUS('SRVRESSM'))

	XSTATUS:=.T.
	@ 9,0 DCSAY  'Enter Sales Person 3 digit Numeric Employee Number'  get cSM
	@ 10,0 DCSAY '           Enter Start Date to Display  / ESC=EXIT'  get dQDate //getfont '8.Courier New' sayfont '.Courier New'
	@ 11,0 DCSAY '                      Enter Ending Date to Display'  get dEnddate VALID{ || IIF( dEnddate < dQdate,.f. ,.t. )}
	dcREAD gui TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA enterexit fit title 'Enter Dates to Display' OPTIONS GETOPTIONS
	IF !XSTATUS
		CLOSE DATABASES
		RETURN NIL
	ENDIF

	SELECT SERVRESV
	SERVRESV->(DBSETFILTER({|| SERVRESV->RESDATE >=dQDate .AND. SERVRESV->RESDATE <=dEnddate}))

	SELECT SERVRESV
	SERVRESV->(DBSEEK(cSm))
	SERVRESV->(DC_SETSCOPE(0,cSm))
	SERVRESV->(DC_SETSCOPE(1,cSm))
	SERVRESV->(DC_DBGOTOP())


	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
	 { XBP_PP_COL_DA_HILITE_FGCLR, GRA_CLR_BLACK }, /*  HILITE FG Color  */     ;
	 { XBP_PP_COL_DA_ROWHEIGHT, 40 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 10 }  }              /* Cell Height */

  aColPres := { { XBPCOL_DA_CELLALIGNMENT, XBPALIGN_RIGHT + XBPALIGN_TOP }, ;
              { XBP_PP_COL_HA_ALIGNMENT, XBPALIGN_RIGHT } }


/* ----- Create ToolBar ----- */

	@ 2,1 DCTOOLBAR ToToolBar                               ;
		SIZE 100, 2 BUTTONSIZE 20,2



	DCADDBUTTONXP CAPTION 'View History'                              ;
		ACTION {|| GNERVHIST(SUBSTR(SERVRESV->VIN,10,8),aChars),DC_GetRefresh(GetList),oBrowse:forcestable()}        ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'VIEW FORWARD SCHEDULE' ;
		CONFIG oConfig2



	DCADDBUTTONXP CAPTION 'Free Form eMail '   ;
		ACTION {|| FREEFORMMAIL(SERVRESV->EMAIL),;
		DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'eMail Functions Free Form'    ;
		CONFIG oConfig2

	 DCADDBUTTONXP CAPTION 'Review Deal Info '   ;
		ACTION {|| _LookCustDealInfo(SERVRESV->LAST8),;
		DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'Look Up Deal'    ;
		CONFIG oConfig2

	 DCADDBUTTONXP CAPTION 'Print This :Browser '   ;
		ACTION {|| _PrintResv(),;
		SERVRESV->(DC_DBGOTOP()),;
		DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'Look Up Deal'    ;
		CONFIG oConfig2




	if SYSCODE() <> 'F'
		DCADDBUTTONXP CAPTION 'Evip'                              ;
			ACTION {|| evip(SERVRESV->VIN,SERVRESV->LASTNAME,.f.,str(SERVCUST->LASTMILES)), DC_GetRefresh(GetList)}        ;
			PARENT ToToolBar                                     ;
			TOOLTIP 'Request Evip'   ;
			BITMAP BITMAP_BROWSER_M      ;
			CONFIG oConfig2
	endif

	IF SYSCODE()='F'
		DCADDBUTTONXP CAPTION 'Y.Oasis'                              ;
			HELPCODE '.OASIS' ;
			ACTION {|| GetVehHist(SERVRESV->VIN,SERVRESV->MILEAGE),DC_GetRefresh(GetList)}        ;
			PARENT toToolBar                                     ;
			TOOLTIP 'Request OASIS';
			BITMAP BITMAP_BROWSER_M      ;
		   CONFIG oConfig2
	ENDIF


/* ----- Create browse ----- */

	@ 4,1 DCBROWSE oBrowse ALIAS 'SERVRESV'                 ;
		SIZE 170,25       ;
		PRESENTATION aPres;
		HEADLINES 4  ;
		FONT '8.Lucida Console'
		//FREEZELEFT {1,2,3};
		//SCOPE

	DCBROWSECOL DATA {|| SERVRESV->SALESMAN} ;
		HEADER "SalesPers" PARENT oBrowse ;
		WIDTH 3

	DCBROWSECOL DATA {|| SERVRESV->RESDATE}                     ;
		HEADER "ApptDate" PARENT oBrowse    ;
		WIDTH 8


	DCBROWSECOL DATA {{|| SERVRESV->FIRSTNAME}, ;
							{|| alltrim(SERVRESV->LASTNAME)+SERVRESV->BUSNAME}};
	  	HEADER "Customer;Name" PARENT oBrowse     ;
	  	WIDTH 20

	DCBROWSECOL DATA {{ || SERVRESV->YEAR}, ;
	                  {|| _GETSERVCUSTINFO(SERVRESV->LAST8,'MAKE')},  ;
							{|| _GETSERVCUSTINFO(SERVRESV->LAST8,'CARDESC')}} ;
		 HEADER 'Year;Make;Model' PARENT oBrowse WIDTH 15

	DCBROWSECOL DATA {{|| _GETSERVCUSTINFO(SERVRESV->LAST8,'lastmiles')},  ;
							{|| _GETSERVCUSTINFO(SERVRESV->LAST8,'lastserv')}} ;
		 HEADER 'LastMiles;LastDteIn' PARENT oBrowse WIDTH 12

	DCBROWSECOL DATA {{|| _GETSERVCUSTINFO(SERVRESV->LAST8,'Deldate')},  ;
							{|| _GETSERVCUSTPAYINFO(SERVRESV->LAST8) }, ;
						   {|| SERVRESV->VIN	}};
		 HEADER 'DelDate;LastPymtDt;Vin' PARENT oBrowse WIDTH 16


	DCBROWSECOL DATA {{|| _GETSERVCUSTINFO(SERVRESV->LAST8,'FINAMT')},  ;
							{|| _GETSERVCUSTINFO(SERVRESV->LAST8,'Term')},;
							{|| _GETSERVCUSTINFO(SERVRESV->LAST8,'Monpymnt')}} ;
		 HEADER '$Financed;Term;$Payment' PARENT oBrowse WIDTH 8

	DCBROWSECOL DATA {{ || SERVRESV->Waityn}, ;
	                  {||  SERVRESV->APPTTIME }}  ;
		 HEADER 'Waiting;Appt Time' PARENT oBrowse WIDTH 8


	DCBROWSECOL DATA {{|| SUBSTR(SERVRESV->CONCERN,1,30)},;
							{|| SUBSTR(SERVRESV->CONCERN,31,30)},;
							{|| SUBSTR(SERVRESV->CONCERN2,1,30)}, ;
							{|| SUBSTR(SERVRESV->CONCERN2,31,30)}}  ;
		HEADER 'Repair:Needs'  PARENT oBrowse WIDTH 20

	DCBROWSECOL DATA {{|| SERVRESV->STREET},;
							{|| SERVRESV->CITYST},;
							{|| SERVRESV->ZIP}} ;
		HEADER 'Address;Info'  PARENT oBrowse WIDTH 15

	DCBROWSECOL FIELD SERVRESV->EMAIL ;
			HEADER "EMAIL" PARENT oBrowse     ;
			WIDTH 20

	DCBROWSECOL DATA {{|| TRANSFORM(SERVRESV->HPHONEALPH,'@R 999.999.9999')},;
							{|| TRANSFORM(SERVRESV->WPHONEALPH,'@R 999.999.9999')},;
							{|| TRANSFORM(SERVRESV->CELL,'@R 999.999.9999')}} ;
		HEADER 'HomePhn;WorkPhn;CellPhn' PARENT oBrowse WIDTH 14


	DCREAD GUI ;
		FIT ;
		BUTTONS DCGUI_BUTTON_EXIT    ;
		TITLE 'Service Appointment Browser' ;
		APPWINDOW MAINWINDOW():DRAWINGAREA;
		OPTIONS GETOPTIONS


   AboCloseAll()
RETURN NIL

static function _PrintResvHdr(nPage)
	nPage++
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Customers with Service Reservations  Page '+ alltrim(str(nPage))
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'SM'
	@ DC_PRINTERROW(),4 DCPRINT SAY   'ApptDate'
	@ DC_PRINTERROW(),15 DCPRINT SAY  'Firstname'
	@ DC_PRINTERROW(),35 DCPRINT SAY  'Year'  FONT '8.Courier New Bold'
	@ DC_PRINTERROW(),50 DCPRINT SAY  'LastMiles'
	@ DC_PRINTERROW(),60 DCPRINT SAY  'Deldate'
	@ DC_PRINTERROW(),70 DCPRINT SAY  '$Financed'
	@ DC_PRINTERROW(),80 DCPRINT SAY  'Wait'
	@ DC_PRINTERROW(),85 DCPRINT SAY 'Street'
	@ DC_PRINTERROW(),105 DCPRINT SAY 'eMail'
	@ DC_PRINTERROW(),130 DCPRINT SAY 'HomePhn'

	@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
	@ DC_PRINTERROW(),4 DCPRINT SAY ''
	@ DC_PRINTERROW(),15 DCPRINT SAY  'LastName'
	@ DC_PRINTERROW(),35 DCPRINT SAY 'Make'
	@ DC_PRINTERROW(),50 DCPRINT SAY 'LastDtIn'
	@ DC_PRINTERROW(),60 DCPRINT SAY 'LastPyDt'
	@ DC_PRINTERROW(),70 DCPRINT SAY 'Term'
	@ DC_PRINTERROW(),80 DCPRINT SAY 'Time'
	@ DC_PRINTERROW(),85 DCPRINT SAY 'City'
	@ DC_PRINTERROW(),130 DCPRINT SAY 'WorkPhn'

	@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
	@ DC_PRINTERROW(),4 DCPRINT SAY ''
	@ DC_PRINTERROW(),15 DCPRINT SAY 'Vin'
	@ DC_PRINTERROW(),35 DCPRINT SAY 'Model'
	@ DC_PRINTERROW(),50 DCPRINT SAY ''
	@ DC_PRINTERROW(),60 DCPRINT SAY ''
	@ DC_PRINTERROW(),70 DCPRINT SAY '$MonPyMnt'
	@ DC_PRINTERROW(),80 DCPRINT SAY ''
	@ DC_PRINTERROW(),85 DCPRINT SAY 'State Zip'
	@ DC_PRINTERROW(),130 DCPRINT SAY 'CellPhn'
	@ DC_PRINTERROW()+1.5,0,DC_PRINTERROW()+1.5,140 DCPRINT LINE

	RETURN NIL


static function _PrintResv()
	LOCAL nOrient:=2 , nDefmode:=4,nCopies:=1 , oPrinter , cFont:='8.Courier New' , nPage:=0 ,cOutfile:='', i
	LOCAL cLast:=''
	TOP_MAR(1)
	BOT_MAR(1)
	PAGE_LEN(60)
	IF !Print_Choice( 'Customer Service Appts', @nCopies,' ',.F.,@nOrient,@cFont,@nDefMode,cOutfile )
		RETURN NIL
	ENDIF
	IF !PrintOn('Customer Service Appts', oPrinter, nOrient, cFont, nCopies)
		  RETURN NIL
	ENDIF
	_PrintResvHdr(@nPage)
	SELECT SERVRESV
	SERVRESV->(DC_DBGOTOP())

	DO WHILE !SERVRESV->(DC_EOF())

	   @ DC_PRINTERROW()+1,0 DCPRINT SAY SERVRESV->SALESMAN
	   @ DC_PRINTERROW(),4 DCPRINT SAY   SERVRESV->RESDATE
	   @ DC_PRINTERROW(),15 DCPRINT SAY  SERVRESV->FIRSTNAME
		@ DC_PRINTERROW(),35 DCPRINT SAY  SERVRESV->Year  FONT '8.Courier New Bold'
		@ DC_PRINTERROW(),50 DCPRINT SAY  _GETSERVCUSTINFO(SERVRESV->LAST8,'lastmiles')
		@ DC_PRINTERROW(),60 DCPRINT SAY  _GETSERVCUSTINFO(SERVRESV->LAST8,'Deldate')
	   @ DC_PRINTERROW(),70 DCPRINT SAY  _GETSERVCUSTINFO(SERVRESV->LAST8,'FINAMT')
		@ DC_PRINTERROW(),80 DCPRINT SAY  SERVRESV->WaitYn
	   @ DC_PRINTERROW(),85 DCPRINT SAY  SERVRESV->Street
	   @ DC_PRINTERROW(),105 DCPRINT SAY SUBSTR(SERVRESV->eMail,1,24)
		@ DC_PRINTERROW(),130 DCPRINT SAY TRANSFORM(SERVRESV->HPHONEALPH,'@R 999.999.9999')


		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY ''
		cLast:=alltrim(SERVRESV->LASTNAME)+SERVRESV->BUSNAME
		@ DC_PRINTERROW(),15 DCPRINT SAY  SUBSTR(clast,1,20)
		@ DC_PRINTERROW(),35 DCPRINT SAY _GETSERVCUSTINFO(SERVRESV->LAST8,'MAKE')
		@ DC_PRINTERROW(),50 DCPRINT SAY _GETSERVCUSTINFO(SERVRESV->LAST8,'lastserv')
		@ DC_PRINTERROW(),60 DCPRINT SAY _GETSERVCUSTPAYINFO(SERVRESV->LAST8)
	   @ DC_PRINTERROW(),70 DCPRINT SAY _GETSERVCUSTINFO(SERVRESV->LAST8,'Term')
		@ DC_PRINTERROW(),80 DCPRINT SAY SERVRESV->APPTTIME
	   @ DC_PRINTERROW(),85 DCPRINT SAY SERVRESV->CityST
		@ DC_PRINTERROW(),105 DCPRINT SAY SUBSTR(SERVRESV->eMail,25,25)
	   @ DC_PRINTERROW(),130 DCPRINT SAY TRANSFORM(SERVRESV->WPHONEALPH,'@R 999.999.9999')

		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY SERVRESV->VIN
		@ DC_PRINTERROW(),15 DCPRINT SAY ''
		@ DC_PRINTERROW(),35 DCPRINT SAY _GETSERVCUSTINFO(SERVRESV->LAST8,'CARDESC')
		@ DC_PRINTERROW(),50 DCPRINT SAY ''
	   @ DC_PRINTERROW(),60 DCPRINT SAY ''
		@ DC_PRINTERROW(),70 DCPRINT SAY _GETSERVCUSTINFO(SERVRESV->LAST8,'Monpymnt')
		@ DC_PRINTERROW(),80 DCPRINT SAY ''
		@ DC_PRINTERROW(),85 DCPRINT SAY SERVRESV->ZIP
	   @ DC_PRINTERROW(),130 DCPRINT SAY TRANSFORM(SERVRESV->CELL,'@R 999.999.9999')
		SKIPALINE()
		IF DCPAGEEJECT()
		  _PrintResvHdr(@nPage)
		ENDIF
		@ DC_PRINTERROW()+1,0,DC_PRINTERROW()+1,140 DCPRINT LINE
		IF DCPAGEEJECT()
		  _PrintResvHdr(@nPage)
		ENDIF

		SERVRESV->(DC_DBSKIP(1))
	ENDDO

	PRINTOFF(oPrinter)
	return nil



function browBIRTHbySM()
	local getlist:={},getoptions,nRow,nCol,aColpres
	LOCAL xStatus, nRec,oConfig1,oConfig2,oDialog, cSM:=space(3),dNextbday,nAge:=0
	LOCAL oBrowse,apres,totoolbar,dQdate:=DATE(), aPush:={}, aChars:={} ,dEnddate:=DATE()+1, aArray:={}
	LOCAL aWaitcolor:={GRA_CLR_BLACK,GRA_CLR_WHITE}
	//SET DELETED ON
	DCGEToptions autoresize saywidth 0 SAYFONT('10.Courier New') getfont('10.Courier New')

	oConfig1 := DC_XbpPushButtonXPConfig():new()
	oConfig1:bitmapOffset := 5
	oConfig1:fgColorMouse := COLOR_BLACK
	oConfig1:bgColorMouse := COLOR_ORANGE
	oConfig1:fgColor := COLOR_BLACK
	oConfig1:bgColor := BD_LINEN
	oConfig1:gradientStep := 10
	oConfig1:gradientReverse := .t.
	oConfig1:radius := 10
	oConfig1:bordercolor := COLOR_RED

	oConfig2 := DC_XbpPushButtonXPConfig():new()
	oConfig2:bitmapOffset := 5
	oConfig2:fgColorMouse := COLOR_BLACK
	oConfig2:bgColorMouse := COLOR_ORANGE
	oConfig2:fgColor := COLOR_BLACK
	oConfig2:bgColor := COLOR_SILVER
	oConfig2:gradientStep := 10
	oConfig2:gradientReverse := .t.
	oConfig2:radius := 10
	oConfig2:bordercolor := COLOR_BLUE



	oDialog:=DC_WAITON('Now Opening Files')
	IF !OPENROFILES()
		DC_IMPL(oDialog)
		RETURN NIL
	ENDIF
	if !OPENSAKHISTFILES()
	   DC_IMPL(oDialog)
	   return nil
   endif
	DC_IMPL(oDialog)

	IF !UseDb({'VTMAS','TEAM'})
	  DC_WINALERT('Cannot Open VTMAS OR TEAM')
	  return nil
   ENDIF




	SERVCUST->(ORDSETFOCUS('SERVSM'))

	XSTATUS:=.T.
	@ 9,0 DCSAY  'Enter Sales Person 3 digit Numeric Employee Number'  get cSM
	@ 10,0 DCSAY '           Enter Start Date to Display  / ESC=EXIT'  get dQDate //getfont '8.Courier New' sayfont '.Courier New'
	@ 11,0 DCSAY '                      Enter Ending Date to Display'  get dEnddate VALID{ || IIF( dEnddate < dQdate,.f. ,.t. )}
	dcREAD gui TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA enterexit fit title 'Enter Dates to Display' OPTIONS GETOPTIONS
	IF !XSTATUS
		CLOSE DATABASES
		RETURN NIL
	ENDIF

	SELECT SERVCUST
	//SERVRESV->(DBSETFILTER({|| _nextbday(SERVCUST->BUYERDOB) >=dQDate .AND. _nextbday(SERVCUST->BUYERDOB) <=dEnddate}))

	SELECT SERVCUST
	SERVCUST->(DBSEEK(cSm))
	SERVCUST->(DC_SETSCOPE(0,cSm))
	SERVCUST->(DC_SETSCOPE(1,cSm))
	SERVCUST->(DC_DBGOTOP())

	DO WHILE !SERVCUST->(DC_EOF())
		dNextbday:=ctod("  /  /  ")
		nAge:=0
		IF _NEXTBDAY(SERVCUST->BUYERDOB,dQdate,dEnddate,@dNextbday,@nAge)

			aADD(aArray,{SERVCUST->SM,ALLTRIM(SERVCUST->LASTNAME)+SERVCUST->BUSNAME,SERVCUST->FIRSTNAME,SERVCUST->BUYERDOB,dNextbday,ROUND(nAge,0),SERVCUST->EMAIL,SERVCUST->STREET,SERVCUST->CITYSTATE,SERVCUST->STATE,SERVCUST->ZIP,;
				SERVCUST->HPHONEALPH,SERVCUST->WPHONEALPH,SERVCUST->CELL,SERVCUST->YEAR,SERVCUST->MAKE,SERVCUST->CARDESC,SERVCUST->LASTMILES,SERVCUST->LASTSERV,SERVCUST->DELDATE,;
				SERVCUST->FINAMT,SERVCUST->TERM,SERVCUST->MONPYMNT,SERVCUST->VIN,SERVCUST->LAST8,SERVCUST->STOCKNO,SERVCUST->(RECNO()),;
				IIF(!EMPTY(SERVCUST->SALECOMM) ,.T. ,.F. )} )
		ENDIF
		SERVCUST->(DBSKIP(1))
	ENDDO
	IF EMPTY(aArray)
		DC_WINALERT('There are no records that fit the criteria entered.')
      AboCloseAll()
		RETURN NIL
	ENDIF
	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
	 { XBP_PP_COL_DA_HILITE_FGCLR, GRA_CLR_BLACK }, /*  HILITE FG Color  */     ;
	 { XBP_PP_COL_DA_ROWHEIGHT, 40 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 10 }  }              /* Cell Height */


 aColPres := { { XBPCOL_DA_CELLALIGNMENT, XBPALIGN_RIGHT + XBPALIGN_TOP }, ;
              { XBP_PP_COL_HA_ALIGNMENT, XBPALIGN_RIGHT } }

/* ----- Create ToolBar ----- */

	@ 2,1 DCTOOLBAR ToToolBar                               ;
		SIZE 100, 2 BUTTONSIZE 20,2


	DCADDBUTTONXP CAPTION 'View History'                              ;
		ACTION {|| GNERVHIST(DC_GetColArray(25,oBrowse),aChars),DC_GetRefresh(GetList),oBrowse:forcestable()}        ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'VIEW FORWARD SCHEDULE' ;
		CONFIG oConfig2




	DCADDBUTTONXP CAPTION 'Free Form eMail '   ;
		ACTION {|| FREEFORMMAIL(DC_GETCOLARRAY(7,oBrowse)),;
		DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'eMail Functions Free Form'    ;
		CONFIG oConfig2

	 DCADDBUTTONXP CAPTION 'Review Deal Info '   ;
		ACTION {|| _LookCustDealInfo(DC_GETCOLARRAY(25,oBrowse)),;
		DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'Look Up Deal'    ;
		CONFIG oConfig2

	 DCADDBUTTONXP CAPTION 'Add Sales ;Follow Up Comments '   ;
		ACTION {|| _AddSalesNotes(DC_GETCOLARRAY(27,oBrowse)),;
		DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'Add Notes to follow up file'    ;
		CONFIG oConfig2

	 DCADDBUTTONXP CAPTION 'Print This ;Browser '   ;
		ACTION {|| _PrintBirth(aArray) };
		PARENT ToToolBar                                     ;
		TOOLTIP 'Add Notes to follow up file'    ;
		CONFIG oConfig2




    aArray:=ASort( aArray,,, {|aX,aY| aX[5] < aY[5]} )



/* ----- Create browse ----- */

	@ 4,1 DCBROWSE oBrowse DATA aArray                 ;
		SIZE 170,25                                     ;
		PRESENTATION aPres;
		HEADLINES 3  ;
		FONT '8.Lucida Console'


		//SCOPE

	DCBROWSECOL ELEMENT 1;
		HEADER "SalesPers" PARENT oBrowse ;
		WIDTH 3

	DCBROWSECOL DATA { {|| Alltrim(DC_GetColArray(3,oBrowse))}, ;
                      {|| SUBSTR(DC_GetColArray(2,oBrowse),1,20)}, ;
                      {|| dtoc(DC_GetColArray(4,oBrowse))} } ;
                      HEADER 'FirstName;LastName;BirthDate' WIDTH 15 PARENT oBrowse

	DCBROWSECOL DATA { {|| dtoc(DC_GetColArray(5,oBrowse))}, ;
                      {|| alltrim(str(DC_GetColArray(6,oBrowse)))}} ;
                      HEADER 'Birthday;Will Be Age' WIDTH 12 PARENT oBrowse HCOLOR 1,BD_KEYLIME;

	DCBROWSECOL DATA { {|| DC_GetColArray(7,oBrowse)}, ;
                      {|| DC_GetColArray(24,oBrowse)},  ;
							 {|| DC_GetColArray(25,oBrowse)}};
                      HEADER 'eMail;Vin;Last8' WIDTH 20 PARENT oBrowse ;

  DCBROWSECOL DATA  {{|| DC_GetColArray(8,oBrowse)}, ;
	                   {|| DC_GetColArray(9,oBrowse)},;
							 {|| DC_GetColArray(10,oBrowse)+' '+DC_GetColArray(11,oBrowse)}} ;
		HEADER "Street;City;State;Zip" PARENT oBrowse     ;
		WIDTH 20

	DCBROWSECOL DATA  {{|| TRANSFORM(DC_GetColArray(12,oBrowse),'@R 999.999.9999')},;
	                   {|| TRANSFORM(DC_GetColArray(13,oBrowse),'@R 999.999.9999')},;
							 {|| TRANSFORM(DC_GetColArray(14,oBrowse),'@R 999.999.9999')}};
		HEADER "HomePhn;WorkPhn;CellPhn" PARENT oBrowse     ;
		WIDTH 14

	DCBROWSECOL DATA  {{|| DC_GetColArray(15,oBrowse)}, ;
	                   {|| DC_GetColArray(16,oBrowse)},;
							 {|| DC_GetColArray(17,oBrowse)}};
		HEADER "Year;Make;Model" PARENT oBrowse     ;
		WIDTH 14

	DCBROWSECOL DATA  {{|| DC_GetColArray(18,oBrowse)}, ;
	                   {|| DC_GetColArray(19,oBrowse)},;
							 {|| _GETSERVCUSTPAYINFO(DC_GetColArray(25,oBrowse))}};
		HEADER "LastMiles;LastServDt;LastPymtDt" PARENT oBrowse     ;
		WIDTH 12

	DCBROWSECOL DATA  {{|| DC_GetColArray(21,oBrowse)}, ;
	                   {|| DC_GetColArray(22,oBrowse)},;
							 {|| DC_GetColArray(23,oBrowse)}};
		HEADER "$Financed;Term;$Payment" PARENT oBrowse     ;
		WIDTH 10


	DCBROWSECOL DATA  {{|| DC_GetColArray(26,oBrowse)}, ;
	                   {|| DC_GetColArray(20,oBrowse)}};
		HEADER "Stock#;Delivery Date" PARENT oBrowse     ;
		WIDTH 15

	DCBROWSECOL DATA {|| IIF( DC_GetColArray(28,oBrowse),'See Comments' ,' ' )};
		 HEADER "Comment;Content"  PARENT oBrowse ;
		 WIDTH 15  COLOR {|| IIF( DC_GetColArray(28,oBrowse) ,{1,BD_KEYLIME} ,{1,BD_LINEN} )}




		DCREAD GUI ;
		FIT ;
		BUTTONS DCGUI_BUTTON_EXIT    ;
		TITLE 'Customer Birthday Browser' ;
		APPWINDOW MAINWINDOW():DRAWINGAREA;
		OPTIONS GETOPTIONS


   AboCloseAll()
RETURN NIL

static function _PrintBirthHdr(nPage)
	nPage++
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Customers with Upcoming Birthdays  Page'+ alltrim(str(nPage))
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'SM'
	@ DC_PRINTERROW(),4 DCPRINT SAY   'FirstName'
	@ DC_PRINTERROW(),20 DCPRINT SAY  'Birthdate'
		@ DC_PRINTERROW(),30 DCPRINT SAY  'eMail'  FONT '8.Courier New Bold'
		@ DC_PRINTERROW(),55 DCPRINT SAY  'Street'
		@ DC_PRINTERROW(),70 DCPRINT SAY  'HomePhn'
	   @ DC_PRINTERROW(),85 DCPRINT SAY  'Year'
		@ DC_PRINTERROW(),100 DCPRINT SAY  'LastMiles'
	   @ DC_PRINTERROW(),115 DCPRINT SAY '$Financed'
	   @ DC_PRINTERROW(),125 DCPRINT SAY 'Stock#'

		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY 'LastName'
		@ DC_PRINTERROW(),20 DCPRINT SAY  'Will Be'
		@ DC_PRINTERROW(),30 DCPRINT SAY 'Vin'
		@ DC_PRINTERROW(),55 DCPRINT SAY 'City'
		@ DC_PRINTERROW(),70 DCPRINT SAY 'WorkPhn'
	   @ DC_PRINTERROW(),85 DCPRINT SAY 'Make'
		@ DC_PRINTERROW(),100 DCPRINT SAY 'LastServ'
	   @ DC_PRINTERROW(),115 DCPRINT SAY 'Term'
	   @ DC_PRINTERROW(),125 DCPRINT SAY 'Delivdate'

		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY 'Birthdate'
		@ DC_PRINTERROW(),30 DCPRINT SAY 'Last8'
		@ DC_PRINTERROW(),55 DCPRINT SAY 'State Zip'
		@ DC_PRINTERROW(),70 DCPRINT SAY 'CellPhn'
	   @ DC_PRINTERROW(),85 DCPRINT SAY 'Model'
		@ DC_PRINTERROW(),100 DCPRINT SAY 'LastPayDt'
	   @ DC_PRINTERROW(),115 DCPRINT SAY 'MonPyMnt'
	   //@ DC_PRINTERROW(),125 DCPRINT SAY 'Last8'
		@ DC_PRINTERROW()+1.5,0,DC_PRINTERROW()+1.5,140 DCPRINT LINE

		RETURN NIL


static function _PrintBirth(aArray)
	LOCAL nOrient:=2 , nDefmode:=4,nCopies:=1 , oPrinter , cFont:='8.Courier New' , nPage:=0 ,cOutfile:='', i

	TOP_MAR(1)
	BOT_MAR(1)
	PAGE_LEN(60)
	IF !Print_Choice( 'Customer Birthdays', @nCopies,' ',.F.,@nOrient,@cFont,@nDefMode,cOutfile )
		RETURN NIL
	ENDIF
	IF !PrintOn('Customer Birthdays', oPrinter, nOrient, cFont, nCopies)
		  RETURN NIL
	ENDIF
	_PrintBirthHdr(@nPage)




	FOR i := 1 TO LEN(aArray)
		@ DC_PRINTERROW()+1,0 DCPRINT SAY aArray[i,1]   //sm
		@ DC_PRINTERROW(),4 DCPRINT SAY aArray[i,3]     //first
		@ DC_PRINTERROW(),20 DCPRINT SAY aArray[i,5]    //nextbdy
		@ DC_PRINTERROW(),30 DCPRINT SAY aArray[i,7]    //email
		@ DC_PRINTERROW(),55 DCPRINT SAY aArray[i,8]    ///street
		@ DC_PRINTERROW(),70 DCPRINT SAY TRANSFORM(aArray[i,12],'@R 999.999.9999')    ///homephn
	   @ DC_PRINTERROW(),85 DCPRINT SAY aArray[i,15]    ///year
		@ DC_PRINTERROW(),100 DCPRINT SAY aArray[i,18]    ///lastmiles
	   @ DC_PRINTERROW(),115 DCPRINT SAY aArray[i,21]    ///$financed
	   @ DC_PRINTERROW(),125 DCPRINT SAY aArray[i,26]    ///stockno
		IF DCPAGEEJECT()
			_PrintBirthHdr(@nPage)
		ENDIF
		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY SUBSTR(aArray[i,2],1,20)     //last
		@ DC_PRINTERROW(),20 DCPRINT SAY aArray[i,6]  PICTURE '999Q'  //newage
		@ DC_PRINTERROW(),30 DCPRINT SAY aArray[i,24]   //vin
		@ DC_PRINTERROW(),55 DCPRINT SAY aArray[i,9]    //city
		@ DC_PRINTERROW(),70 DCPRINT SAY TRANSFORM(aArray[i,13],'@R 999.999.9999')    ///workphn
	   @ DC_PRINTERROW(),85 DCPRINT SAY aArray[i,16]    ///make
		@ DC_PRINTERROW(),100 DCPRINT SAY aArray[i,19]    ///lastserv
	   @ DC_PRINTERROW(),115 DCPRINT SAY aArray[i,22]    ///term
	   @ DC_PRINTERROW(),125 DCPRINT SAY aArray[i,20]    ///deldate
		IF DCPAGEEJECT()
			_PrintBirthHdr(@nPage)
		ENDIF

		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY aArray[i,4]     //dob
		//@ DC_PRINTERROW(),20 DCPRINT SAY aArray[i,25] //last8
		@ DC_PRINTERROW(),30 DCPRINT SAY aArray[i,25]
		@ DC_PRINTERROW(),55 DCPRINT SAY aArray[i,10]+' '+aArray[i,11]    ///state zip
		@ DC_PRINTERROW(),70 DCPRINT SAY TRANSFORM(aArray[i,14],'@R 999.999.9999')    ///cellphn
	   @ DC_PRINTERROW(),85 DCPRINT SAY aArray[i,17]    ///model
		@ DC_PRINTERROW(),100 DCPRINT SAY _GETSERVCUSTPAYINFO(aArray[i,25])    ///lastpay
	   @ DC_PRINTERROW(),115 DCPRINT SAY aArray[i,23]    ///monpymt
		IF DCPAGEEJECT()
			_PrintBirthHdr(@nPage)
		ENDIF
		SKIPALINE()
		IF DCPAGEEJECT()
			_PrintBirthHdr(@nPage)
		ENDIF
		@ DC_PRINTERROW()+1,0,DC_PRINTERROW()+1,140 DCPRINT LINE
		IF DCPAGEEJECT()
			_PrintBirthHdr(@nPage)
		ENDIF
	NEXT

	PRINTOFF(oPrinter)
	return nil







function browRecentInbySM()
	local getlist:={},getoptions,nRow,nCol,aColpres
	LOCAL xStatus, nRec,oConfig1,oConfig2,oDialog, cSM:=space(3),dNextbday,nAge:=0
	LOCAL oBrowse,apres,totoolbar,dQdate:=DATE(), aPush:={}, aChars:={} ,dEnddate:=DATE()+1, aArray:={}
	LOCAL aWaitcolor:={GRA_CLR_BLACK,GRA_CLR_WHITE}
	//SET DELETED ON
	DCGEToptions autoresize saywidth 0 SAYFONT('10.Courier New') getfont('10.Courier New')

	oConfig1 := DC_XbpPushButtonXPConfig():new()
	oConfig1:bitmapOffset := 5
	oConfig1:fgColorMouse := COLOR_BLACK
	oConfig1:bgColorMouse := COLOR_ORANGE
	oConfig1:fgColor := COLOR_BLACK
	oConfig1:bgColor := BD_LINEN
	oConfig1:gradientStep := 10
	oConfig1:gradientReverse := .t.
	oConfig1:radius := 10
	oConfig1:bordercolor := COLOR_RED

	oConfig2 := DC_XbpPushButtonXPConfig():new()
	oConfig2:bitmapOffset := 5
	oConfig2:fgColorMouse := COLOR_BLACK
	oConfig2:bgColorMouse := COLOR_ORANGE
	oConfig2:fgColor := COLOR_BLACK
	oConfig2:bgColor := COLOR_SILVER
	oConfig2:gradientStep := 10
	oConfig2:gradientReverse := .t.
	oConfig2:radius := 10
	oConfig2:bordercolor := COLOR_BLUE



	oDialog:=DC_WAITON('Now Opening Files')
	IF !OPENROFILES()
		DC_IMPL(oDialog)
		RETURN NIL
	ENDIF
	if !OPENSAKHISTFILES()
	   DC_IMPL(oDialog)
	   return nil
   endif
	DC_IMPL(oDialog)

	IF !UseDb({'VTMAS','TEAM'})
	  DC_WINALERT('Cannot Open VTMAS OR TEAM')
	  return nil
   ENDIF




	SERVCUST->(ORDSETFOCUS('SERVSM'))

	XSTATUS:=.T.
	@ 9,0 DCSAY  'Enter Sales Person 3 digit Numeric Employee Number'  get cSM
	@ 10,0 DCSAY '           Enter Start Date to Display  / ESC=EXIT'  get dQDate //getfont '8.Courier New' sayfont '.Courier New'
	@ 11,0 DCSAY '                      Enter Ending Date to Display'  get dEnddate VALID{ || IIF( dEnddate < dQdate,.f. ,.t. )}
	dcREAD gui TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA enterexit fit title 'Enter Dates to Display' OPTIONS GETOPTIONS
	IF !XSTATUS
		CLOSE DATABASES
		RETURN NIL
	ENDIF

	SELECT SERVCUST
	//SERVRESV->(DBSETFILTER({|| _nextbday(SERVCUST->BUYERDOB) >=dQDate .AND. _nextbday(SERVCUST->BUYERDOB) <=dEnddate}))

	SELECT SERVCUST
	SERVCUST->(DBSEEK(cSm))
	SERVCUST->(DC_SETSCOPE(0,cSm))
	SERVCUST->(DC_SETSCOPE(1,cSm))
	SERVCUST->(DC_DBGOTOP())

	DO WHILE !SERVCUST->(DC_EOF())
		IF SERVCUST->LASTSERV >=dQDate .AND. SERVCUST->LASTSERV <=dEnddate
			aADD(aArray,{SERVCUST->SM,ALLTRIM(SERVCUST->LASTNAME)+SERVCUST->BUSNAME,SERVCUST->FIRSTNAME,SERVCUST->BUYERDOB,SERVCUST->EMAIL,SERVCUST->STREET,SERVCUST->CITYSTATE,SERVCUST->STATE,SERVCUST->ZIP,;
				SERVCUST->HPHONEALPH,SERVCUST->WPHONEALPH,SERVCUST->CELL,SERVCUST->YEAR,SERVCUST->MAKE,SERVCUST->CARDESC,SERVCUST->LASTMILES,SERVCUST->LASTSERV,SERVCUST->DELDATE,;
				SERVCUST->FINAMT,SERVCUST->TERM,SERVCUST->MONPYMNT,SERVCUST->VIN,SERVCUST->LAST8,SERVCUST->STOCKNO,SERVCUST->(RECNO()),;
				IIF(!EMPTY(SERVCUST->SALECOMM) ,.T. ,.F. )} )
		ENDIF
		SERVCUST->(DBSKIP(1))
	ENDDO
	IF EMPTY(aArray)
		DC_WINALERT('There are no records that fit the criteria entered.')
      AboCloseAll()
		RETURN NIL
	ENDIF

	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
	 { XBP_PP_COL_DA_HILITE_FGCLR, GRA_CLR_BLACK }, /*  HILITE FG Color  */     ;
	 { XBP_PP_COL_DA_ROWHEIGHT, 40 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 10 }  }              /* Cell Height */


 aColPres := { { XBPCOL_DA_CELLALIGNMENT, XBPALIGN_RIGHT + XBPALIGN_TOP }, ;
              { XBP_PP_COL_HA_ALIGNMENT, XBPALIGN_RIGHT } }

/* ----- Create ToolBar ----- */

	@ 2,1 DCTOOLBAR ToToolBar                               ;
		SIZE 100, 2 BUTTONSIZE 20,2



	DCADDBUTTONXP CAPTION 'View History'                              ;
		ACTION {|| GNERVHIST(DC_GetColArray(23,oBrowse),aChars),DC_GetRefresh(GetList),oBrowse:forcestable()}        ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'VIEW FORWARD SCHEDULE' ;
		CONFIG oConfig2


	DCADDBUTTONXP CAPTION 'Free Form eMail '   ;
		ACTION {|| FREEFORMMAIL(DC_GETCOLARRAY(5,oBrowse)),;
		DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'eMail Functions Free Form'    ;
		CONFIG oConfig2

	 DCADDBUTTONXP CAPTION 'Review Deal Info '   ;
		ACTION {|| _LookCustDealInfo(DC_GETCOLARRAY(23,oBrowse)),;
		DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'Look Up Deal'    ;
		CONFIG oConfig2

	 DCADDBUTTONXP CAPTION 'Add Sales ;Follow Up Comments '   ;
		ACTION {|| _AddSalesNotes(DC_GETCOLARRAY(25,oBrowse)),;
		DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'Add Notes to follow up file'    ;
		CONFIG oConfig2

	DCADDBUTTONXP CAPTION 'Print This ;Browser '   ;
		ACTION {||_PrintRecent(aArray) };
		PARENT ToToolBar                                     ;
		TOOLTIP 'Add Notes to follow up file'    ;
		CONFIG oConfig2




    aArray:=ASort( aArray,,, {|aX,aY| aX[5] < aY[5]} )



/* ----- Create browse ----- */

	@ 4,1 DCBROWSE oBrowse DATA aArray                 ;
		SIZE 170,25                                     ;
		PRESENTATION aPres;
		HEADLINES 3  ;
		FONT '8.Lucida Console'


		//SCOPE

	DCBROWSECOL ELEMENT 1;
		HEADER "SalesPers" PARENT oBrowse ;
		WIDTH 3

	DCBROWSECOL DATA { {|| Alltrim(DC_GetColArray(3,oBrowse))}, ;
                      {|| SUBSTR(DC_GetColArray(2,oBrowse),1,20)}, ;
                      {|| dtoc(DC_GetColArray(4,oBrowse))} } ;
                      HEADER 'FirstName;LastName;BirthDate' WIDTH 15 PARENT oBrowse


	DCBROWSECOL DATA { {|| DC_GetColArray(5,oBrowse)}, ;
                      {|| DC_GetColArray(22,oBrowse)},  ;
							 {|| DC_GetColArray(23,oBrowse)}};
                      HEADER 'eMail;Vin;Last8' WIDTH 20 PARENT oBrowse ;

  DCBROWSECOL DATA  {{|| DC_GetColArray(6,oBrowse)}, ;
	                   {|| DC_GetColArray(7,oBrowse)},;
							 {|| DC_GetColArray(8,oBrowse)+' '+DC_GetColArray(9,oBrowse)}} ;
		HEADER "Street;City;State;Zip" PARENT oBrowse     ;
		WIDTH 20


	DCBROWSECOL DATA   {{|| TRANSFORM(DC_GetColArray(10,oBrowse),'@R 999.999.9999')},;
	                   {|| TRANSFORM(DC_GetColArray(11,oBrowse),'@R 999.999.9999')},;
							 {|| TRANSFORM(DC_GetColArray(12,oBrowse),'@R 999.999.9999')}};
		HEADER "HomePhn;WorkPhn;CellPhn" PARENT oBrowse     ;
		WIDTH 15

	DCBROWSECOL DATA  {{|| DC_GetColArray(13,oBrowse)}, ;
	                   {|| DC_GetColArray(14,oBrowse)},;
							 {|| DC_GetColArray(15,oBrowse)}};
		HEADER "Year;Make;Model" PARENT oBrowse     ;
		WIDTH 14

	DCBROWSECOL DATA  {{|| DC_GetColArray(16,oBrowse)}, ;
	                   {|| DC_GetColArray(17,oBrowse)},;
							 {|| _GETSERVCUSTPAYINFO(DC_GetColArray(23,oBrowse))}};
		HEADER "LastMiles;LastServDt;LastPymtDt" PARENT oBrowse     ;
		WIDTH 10

	DCBROWSECOL DATA  {{|| DC_GetColArray(19,oBrowse)}, ;
	                   {|| DC_GetColArray(20,oBrowse)},;
							 {|| DC_GetColArray(21,oBrowse)}};
		HEADER "$Financed;Term;$Payment" PARENT oBrowse     ;
		WIDTH 10


	DCBROWSECOL DATA  {{|| DC_GetColArray(24,oBrowse)}, ;
	                   {|| DC_GetColArray(18,oBrowse)}};
		HEADER "Stock#;Delivery Date" PARENT oBrowse     ;
		WIDTH 15

	DCBROWSECOL DATA {|| IIF( DC_GetColArray(26,oBrowse),'See Comments' ,' ' )};
		 HEADER "Comment;Content"  PARENT oBrowse ;
		 WIDTH 15 COLOR {|| IIF(DC_GetColArray(26,oBrowse) ,{1,BD_KEYLIME} ,{1,BD_LINEN} )}



		DCREAD GUI ;
		FIT ;
		BUTTONS DCGUI_BUTTON_EXIT    ;
		TITLE 'Customer Recently In For Service' ;
		APPWINDOW MAINWINDOW():DRAWINGAREA;
		OPTIONS GETOPTIONS


   AboCloseAll()
RETURN NIL

static function _PrintRecentHdr(nPage)
	nPage++
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Customers Recently In for Service  Page'+ alltrim(str(nPage))
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'SM'
	@ DC_PRINTERROW(),4 DCPRINT SAY   'FirstName'
		@ DC_PRINTERROW(),20 DCPRINT SAY  'eMail'  FONT '8.Courier New Bold'
		@ DC_PRINTERROW(),45 DCPRINT SAY  'Street'
		@ DC_PRINTERROW(),65 DCPRINT SAY  'HomePhn'
	   @ DC_PRINTERROW(),85 DCPRINT SAY  'Year'
		@ DC_PRINTERROW(),100 DCPRINT SAY  'LastMiles'
	   @ DC_PRINTERROW(),115 DCPRINT SAY '$Financed'
	   @ DC_PRINTERROW(),125 DCPRINT SAY 'Stock#'

		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY 'LastName'
		@ DC_PRINTERROW(),20 DCPRINT SAY 'Vin'
		@ DC_PRINTERROW(),45 DCPRINT SAY 'City'
		@ DC_PRINTERROW(),65 DCPRINT SAY 'WorkPhn'
	   @ DC_PRINTERROW(),85 DCPRINT SAY 'Make'
		@ DC_PRINTERROW(),100 DCPRINT SAY 'LastServ'
	   @ DC_PRINTERROW(),115 DCPRINT SAY 'Term'
	   @ DC_PRINTERROW(),125 DCPRINT SAY 'Delivdate'

		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY 'Birthdate'
		@ DC_PRINTERROW(),20 DCPRINT SAY 'Last8'
		@ DC_PRINTERROW(),45 DCPRINT SAY 'State Zip'
		@ DC_PRINTERROW(),65 DCPRINT SAY 'CellPhn'
	   @ DC_PRINTERROW(),85 DCPRINT SAY 'Model'
		@ DC_PRINTERROW(),100 DCPRINT SAY 'LastPayDt'
	   @ DC_PRINTERROW(),115 DCPRINT SAY 'MonPyMnt'
	   //@ DC_PRINTERROW(),125 DCPRINT SAY 'Last8'
		@ DC_PRINTERROW()+1.5,0,DC_PRINTERROW()+1.5,140 DCPRINT LINE

		RETURN NIL


static function _PrintRecent(aArray)
	LOCAL nOrient:=2 , nDefmode:=4,nCopies:=1 , oPrinter , cFont:='8.Courier New' , nPage:=0 ,cOutfile:='', i

	TOP_MAR(1)
	BOT_MAR(1)
	PAGE_LEN(60)
	IF !Print_Choice( 'Recent Customer Visits', @nCopies,' ',.F.,@nOrient,@cFont,@nDefMode,cOutfile )
		RETURN NIL
	ENDIF
	IF !PrintOn('Recent Customer Visits', oPrinter, nOrient, cFont, nCopies)
		  RETURN NIL
	ENDIF
	_PrintRecentHdr(@nPage)

	FOR i := 1 TO LEN(aArray)
		@ DC_PRINTERROW()+1,0 DCPRINT SAY aArray[i,1]   //sm
		@ DC_PRINTERROW(),4 DCPRINT SAY aArray[i,3]     //first
		@ DC_PRINTERROW(),20 DCPRINT SAY aArray[i,5]    //email
		@ DC_PRINTERROW(),45 DCPRINT SAY aArray[i,6]    ///street
		@ DC_PRINTERROW(),65 DCPRINT SAY TRANSFORM(aArray[i,10],'@R 999.999.9999')    ///homephn
	   @ DC_PRINTERROW(),85 DCPRINT SAY aArray[i,13]    ///year
		@ DC_PRINTERROW(),100 DCPRINT SAY aArray[i,16]    ///lastmiles
	   @ DC_PRINTERROW(),115 DCPRINT SAY aArray[i,19]    ///$financed
	   @ DC_PRINTERROW(),125 DCPRINT SAY aArray[i,24]    ///stockno
		IF DCPAGEEJECT()
			_PrintRecentHdr(@nPage)
		ENDIF
		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY aArray[i,2]     //last
		@ DC_PRINTERROW(),20 DCPRINT SAY aArray[i,22]    //vin
		@ DC_PRINTERROW(),45 DCPRINT SAY aArray[i,7]    //city
		@ DC_PRINTERROW(),65 DCPRINT SAY TRANSFORM(aArray[i,11],'@R 999.999.9999')    ///workphn
	   @ DC_PRINTERROW(),85 DCPRINT SAY aArray[i,14]    ///make
		@ DC_PRINTERROW(),100 DCPRINT SAY aArray[i,17]    ///lastserv
	   @ DC_PRINTERROW(),115 DCPRINT SAY aArray[i,20]    ///term
	   @ DC_PRINTERROW(),125 DCPRINT SAY aArray[i,18]    ///deldate
		IF DCPAGEEJECT()
			_PrintRecentHdr(@nPage)
		ENDIF

		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY aArray[i,4]     //dob
		@ DC_PRINTERROW(),20 DCPRINT SAY aArray[i,23] //last8
		@ DC_PRINTERROW(),45 DCPRINT SAY aArray[i,8]+' '+aArray[i,9]    ///state zip
		@ DC_PRINTERROW(),65 DCPRINT SAY TRANSFORM(aArray[i,12],'@R 999.999.9999')    ///cellphn
	   @ DC_PRINTERROW(),85 DCPRINT SAY aArray[i,15]    ///model
		@ DC_PRINTERROW(),100 DCPRINT SAY _GETSERVCUSTPAYINFO(aArray[i,23])    ///lastpay
	   @ DC_PRINTERROW(),115 DCPRINT SAY aArray[i,21]    ///monpymt
		IF DCPAGEEJECT()
			_PrintRecentHdr(@nPage)
		ENDIF
		SKIPALINE()
		IF DCPAGEEJECT()
			_PrintRecentHdr(@nPage)
		ENDIF
		@ DC_PRINTERROW()+1,0,DC_PRINTERROW()+1,140 DCPRINT LINE
		IF DCPAGEEJECT()
			_PrintRecentHdr(@nPage)
		ENDIF
	NEXT

	PRINTOFF(oPrinter)
	return nil




//// function to create browse of upcoming expiring lease or retail payments
function browUpcomingbySM()
	local getlist:={},getoptions,nRow,nCol,aColpres
	LOCAL xStatus, nRec,oConfig1,oConfig2,oDialog, cSM:=space(3),dExpdate
	LOCAL oBrowse,apres,totoolbar,dQdate:=DATE(), aPush:={}, aChars:={} ,dEnddate:=DATE()+1, aArray:={}
	LOCAL aWaitcolor:={GRA_CLR_BLACK,GRA_CLR_WHITE}
	//SET DELETED ON
	DCGEToptions autoresize saywidth 0 SAYFONT('10.Courier New') getfont('10.Courier New')

	oConfig1 := DC_XbpPushButtonXPConfig():new()
	oConfig1:bitmapOffset := 5
	oConfig1:fgColorMouse := COLOR_BLACK
	oConfig1:bgColorMouse := COLOR_ORANGE
	oConfig1:fgColor := COLOR_BLACK
	oConfig1:bgColor := BD_LINEN
	oConfig1:gradientStep := 10
	oConfig1:gradientReverse := .t.
	oConfig1:radius := 10
	oConfig1:bordercolor := COLOR_RED

	oConfig2 := DC_XbpPushButtonXPConfig():new()
	oConfig2:bitmapOffset := 5
	oConfig2:fgColorMouse := COLOR_BLACK
	oConfig2:bgColorMouse := COLOR_ORANGE
	oConfig2:fgColor := COLOR_BLACK
	oConfig2:bgColor := COLOR_SILVER
	oConfig2:gradientStep := 10
	oConfig2:gradientReverse := .t.
	oConfig2:radius := 10
	oConfig2:bordercolor := COLOR_BLUE



	oDialog:=DC_WAITON('Now Opening Files')
	IF !OPENROFILES()
		DC_IMPL(oDialog)
		RETURN NIL
	ENDIF
	if !OPENSAKHISTFILES()
	   DC_IMPL(oDialog)
	   return nil
   endif
	DC_IMPL(oDialog)

	IF !UseDb({'VTMAS','TEAM'})
	  DC_WINALERT('Cannot Open VTMAS OR TEAM')
	  return nil
   ENDIF




	SERVCUST->(ORDSETFOCUS('SERVSM'))

	XSTATUS:=.T.
	@ 7,0 DCSAY  'This function will create a browse of soon to end payment schedules for customers.'
	@ 9,0 DCSAY  'Enter Sales Person 3 digit Numeric Employee Number'  get cSM
	@ 10,0 DCSAY '           Enter Start Date to Display  / ESC=EXIT'  get dQDate //getfont '8.Courier New' sayfont '.Courier New'
	@ 11,0 DCSAY '                      Enter Ending Date to Display'  get dEnddate  VALID{ || IIF( dEnddate < dQdate,.f. ,.t. )}
	dcREAD gui TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA enterexit fit title 'Enter Dates to Display' OPTIONS GETOPTIONS
	IF !XSTATUS
		CLOSE DATABASES
		RETURN NIL
	ENDIF

	SELECT SERVCUST

	SELECT SERVCUST
	SERVCUST->(DBSEEK(cSm))
	SERVCUST->(DC_SETSCOPE(0,cSm))
	SERVCUST->(DC_SETSCOPE(1,cSm))
	SERVCUST->(DC_DBGOTOP())

	DO WHILE !SERVCUST->(DC_EOF())
		IF _FALLSINRANGE(dQdate,dEnddate,@dExpdate)
			aADD(aArray,{SERVCUST->SM,ALLTRIM(SERVCUST->LASTNAME)+SERVCUST->BUSNAME,SERVCUST->FIRSTNAME,SERVCUST->DELDATE,dExpdate,SERVCUST->FINAMT,SERVCUST->TERM,;
				SERVCUST->MONPYMNT,SERVCUST->YEAR,SERVCUST->MAKE,SERVCUST->CARDESC,;
				SERVCUST->EMAIL,SERVCUST->STREET,SERVCUST->CITYSTATE,SERVCUST->STATE,SERVCUST->ZIP,;
				SERVCUST->HPHONEALPH,SERVCUST->WPHONEALPH,SERVCUST->CELL,SERVCUST->LASTMILES,SERVCUST->LASTSERV,;
				SERVCUST->VIN,SERVCUST->LAST8,SERVCUST->STOCKNO,SERVCUST->(RECNO()),IIF(!EMPTY(SERVCUST->SALECOMM) ,.T. ,.F. )} )
		ENDIF
		SERVCUST->(DBSKIP(1))
	ENDDO
	IF EMPTY(aArray)
		DC_WINALERT('There are no records that fit the criteria entered.')
      AboCloseAll()
		RETURN NIL
	ENDIF

	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
	 { XBP_PP_COL_DA_HILITE_FGCLR, GRA_CLR_BLACK }, /*  HILITE FG Color  */     ;
	 { XBP_PP_COL_DA_ROWHEIGHT, 40 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 10 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

	@ 2,1 DCTOOLBAR ToToolBar                               ;
		SIZE 100, 2 BUTTONSIZE 20,2


	DCADDBUTTONXP CAPTION 'View History'                              ;
		ACTION {|| GNERVHIST(DC_GetColArray(23,oBrowse),aChars),DC_GetRefresh(GetList),oBrowse:forcestable()}        ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'VIEW FORWARD SCHEDULE' ;
		CONFIG oConfig2



	DCADDBUTTONXP CAPTION 'Free Form eMail '   ;
		ACTION {|| FREEFORMMAIL(DC_GETCOLARRAY(12,oBrowse)),;
		DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'eMail Functions Free Form'    ;
		CONFIG oConfig2

	 DCADDBUTTONXP CAPTION 'Review Deal Info '   ;
		ACTION {|| _LookCustDealInfo(DC_GETCOLARRAY(23,oBrowse)),;
		DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'Look Up Deal'    ;
		CONFIG oConfig2

	DCADDBUTTONXP CAPTION 'Add Sales ;Follow Up Comments '   ;
		ACTION {|| _AddSalesNotes(DC_GETCOLARRAY(25,oBrowse)),;
		DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'Add Notes to follow up file'    ;
		CONFIG oConfig2

	DCADDBUTTONXP CAPTION 'Print ; Browser '   ;
		ACTION {||_PrintPayUps(aArray)} ;
		PARENT ToToolBar                                     ;
		TOOLTIP 'Print This Browse'    ;
		CONFIG oConfig2




    aArray:=ASort( aArray,,, {|aX,aY| aX[5] < aY[5]} )

	 aColPres := { { XBPCOL_DA_CELLALIGNMENT, XBPALIGN_RIGHT + XBPALIGN_TOP }, ;
              { XBP_PP_COL_HA_ALIGNMENT, XBPALIGN_RIGHT } }



/* ----- Create browse ----- */

	@ 4,1 DCBROWSE oBrowse DATA aArray                 ;
		SIZE 170,25                                     ;
		PRESENTATION aPres;
		HEADLINES 3 ;
	   FONT '8.Lucida Console'

		//FREEZELEFT {1,2,3};
		//SCOPE

	DCBROWSECOL ELEMENT 1;
		HEADER "SalesPers" PARENT oBrowse ;
		WIDTH 3

	DCBROWSECOL DATA { {|| Alltrim(DC_GetColArray(3,oBrowse))}, ;
                      {|| alltrim(DC_GetColArray(2,oBrowse))}, ;
                      {|| IIF( valtype(DC_GetColArray(4,oBrowse))='D',dtoc(DC_GetColArray(4,oBrowse)) ,'' )} } ;
                      HEADER 'FirstName;LastName;DelivDate' WIDTH 15 PARENT oBrowse ;


   DCBROWSECOL ELEMENT 5 ;
		HEADER "Loan Finish" PARENT oBrowse HCOLOR 1,BD_KEYLIME    ;
		WIDTH 5



	DCBROWSECOL DATA  {{|| DC_GetColArray(6,oBrowse)}, ;
	                   {|| DC_GetColArray(7,oBrowse)},;
							 {|| DC_GetColArray(8,oBrowse)}};
		HEADER "$Financed;Term;$Payment" PARENT oBrowse     ;
		WIDTH 10

	DCBROWSECOL DATA  {{|| DC_GetColArray(9,oBrowse)}, ;
	                   {|| DC_GetColArray(10,oBrowse)},;
							 {|| DC_GetColArray(11,oBrowse)}};
		HEADER "Year;Make;Model" PARENT oBrowse     ;
		WIDTH 15

	DCBROWSECOL DATA  {{|| DC_GetColArray(12,oBrowse)},;
							{|| DC_GetColArray(22,oBrowse)},;
							{|| DC_GetColArray(24,oBrowse)}} ;
		HEADER "eMail;Vin;Stock#" PARENT oBrowse     ;
		WIDTH 20

	DCBROWSECOL DATA  {{|| DC_GetColArray(13,oBrowse)}, ;
	                   {|| DC_GetColArray(14,oBrowse)},;
							 {|| DC_GetColArray(15,oBrowse)+' '+DC_GetColArray(16,oBrowse)}} ;
		HEADER "Street;City;State;Zip" PARENT oBrowse     ;
		WIDTH 20


	DCBROWSECOL DATA  {{|| TRANSFORM(DC_GetColArray(17,oBrowse),'@R 999.999.9999')},;
	                   {|| TRANSFORM(DC_GetColArray(18,oBrowse),'@R 999.999.9999')},;
							 {|| TRANSFORM(DC_GetColArray(19,oBrowse),'@R 999.999.9999')}};
		HEADER "HomePhn;WorkPhn;CellPhn" PARENT oBrowse     ;
		WIDTH 14   ;


	DCBROWSECOL DATA  {{|| DC_GetColArray(20,oBrowse)}, ;
	                   {|| DC_GetColArray(21,oBrowse)},;
							 {|| DC_GetColArray(23,oBrowse)}};
		HEADER "LastMiles;LastServ;Last8" PARENT oBrowse     ;
		WIDTH 10

	DCBROWSECOL DATA {|| IIF( DC_GetColArray(26,oBrowse),'See Comments' ,' ' )};
		 HEADER "Comment;Content"  PARENT oBrowse ;
		 WIDTH 15 COLOR {|| IIF(DC_GetColArray(26,oBrowse) ,{1,BD_KEYLIME} ,{1,BD_LINEN} )}


		DCREAD GUI ;
		FIT ;
		BUTTONS DCGUI_BUTTON_EXIT    ;
		TITLE 'Customer Payments Up Browser' ;
		APPWINDOW MAINWINDOW():DRAWINGAREA;
		OPTIONS GETOPTIONS


   AboCloseAll()
RETURN NIL

static function _PrintPayUpsHdr(nPage)
	nPage++
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Customers with Upcoming Payment Expirations  Page'+ alltrim(str(nPage))
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'SM'
	@ DC_PRINTERROW(),4 DCPRINT SAY   'FirstName'
		@ DC_PRINTERROW(),20 DCPRINT SAY  'PayEndDt'  FONT '8.Courier New Bold'
		@ DC_PRINTERROW(),30 DCPRINT SAY  '$Finance'
		@ DC_PRINTERROW(),40 DCPRINT SAY  'Year'
	   @ DC_PRINTERROW(),55 DCPRINT SAY  'eMail'
		@ DC_PRINTERROW(),85 DCPRINT SAY  'Street'
	   @ DC_PRINTERROW(),105 DCPRINT SAY 'HomePhn'
	   @ DC_PRINTERROW(),125 DCPRINT SAY 'Lastmiles'

		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY 'LastName'
		@ DC_PRINTERROW(),30 DCPRINT SAY 'Term'
		@ DC_PRINTERROW(),40 DCPRINT SAY 'Make'
	   @ DC_PRINTERROW(),55 DCPRINT SAY 'Vin'
		@ DC_PRINTERROW(),85 DCPRINT SAY 'City'
	   @ DC_PRINTERROW(),105 DCPRINT SAY 'Workphn'
	   @ DC_PRINTERROW(),125 DCPRINT SAY 'LastServ'

		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY 'Delivdate'
		@ DC_PRINTERROW(),30 DCPRINT SAY'$Payment'
		@ DC_PRINTERROW(),40 DCPRINT SAY 'Model'
	   @ DC_PRINTERROW(),55 DCPRINT SAY 'Stock#'
		@ DC_PRINTERROW(),85 DCPRINT SAY 'State Zip'
	   @ DC_PRINTERROW(),105 DCPRINT SAY 'Cellphn'
	   @ DC_PRINTERROW(),125 DCPRINT SAY 'Last8'
		@ DC_PRINTERROW()+1.5,0,DC_PRINTERROW()+1.5,140 DCPRINT LINE

		RETURN NIL


static function _PrintPayUps(aArray)
	LOCAL nOrient:=2 , nDefmode:=4,nCopies:=1 , oPrinter , cFont:='8.Courier New' , nPage:=0 ,cOutfile:='', i

	TOP_MAR(1)
	BOT_MAR(1)
	PAGE_LEN(60)
	IF !Print_Choice( 'Customer Payments Up', @nCopies,' ',.F.,@nOrient,@cFont,@nDefMode,cOutfile )
		RETURN NIL
	ENDIF
	IF !PrintOn('Payments Up', oPrinter, nOrient, cFont, nCopies)
		  RETURN NIL
	ENDIF
	_PrintPayUpsHdr(@nPage)

	FOR i := 1 TO LEN(aArray)
		@ DC_PRINTERROW()+1,0 DCPRINT SAY aArray[i,1]   //sm
		@ DC_PRINTERROW(),4 DCPRINT SAY aArray[i,3]     //first
		@ DC_PRINTERROW(),20 DCPRINT SAY aArray[i,5] FONT '8.Courier New Bold'    //loanfin
		@ DC_PRINTERROW(),30 DCPRINT SAY aArray[i,6]    ///$fin
		@ DC_PRINTERROW(),40 DCPRINT SAY aArray[i,9]    ///year
	   @ DC_PRINTERROW(),55 DCPRINT SAY aArray[i,12]    ///email
		@ DC_PRINTERROW(),85 DCPRINT SAY aArray[i,13]    ///street
	   @ DC_PRINTERROW(),105 DCPRINT SAY TRANSFORM(aArray[i,17],'@R 999.999.9999')    ///homephn
	   @ DC_PRINTERROW(),125 DCPRINT SAY aArray[i,20]    ///lastmiles
		IF DCPAGEEJECT()
			_PrintPayUpsHdr(@nPage)
		ENDIF
		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY aArray[i,2]     //last
		@ DC_PRINTERROW(),30 DCPRINT SAY aArray[i,7]    //term
		@ DC_PRINTERROW(),40 DCPRINT SAY aArray[i,10]    ///make
	   @ DC_PRINTERROW(),55 DCPRINT SAY aArray[i,22]    ///vin
		@ DC_PRINTERROW(),85 DCPRINT SAY aArray[i,14]    ///city
	   @ DC_PRINTERROW(),105 DCPRINT SAY TRANSFORM(aArray[i,18],'@R 999.999.9999')    ///workphn
	   @ DC_PRINTERROW(),125 DCPRINT SAY aArray[i,21]    ///lastserv
		IF DCPAGEEJECT()
			_PrintPayUpsHdr(@nPage)
		ENDIF

		@ DC_PRINTERROW()+1,0 DCPRINT SAY ''   //
		@ DC_PRINTERROW(),4 DCPRINT SAY aArray[i,4]     //delivdate
		@ DC_PRINTERROW(),30 DCPRINT SAY aArray[i,8]    ///$payment
		@ DC_PRINTERROW(),40 DCPRINT SAY aArray[i,11]    ///model
	   @ DC_PRINTERROW(),55 DCPRINT SAY aArray[i,24]    ///stock
		@ DC_PRINTERROW(),85 DCPRINT SAY aArray[i,15]+' '+aArray[i,16]    ///state zip
	   @ DC_PRINTERROW(),105 DCPRINT SAY TRANSFORM(aArray[i,19],'@R 999.999.9999')    ///cellphn
	   @ DC_PRINTERROW(),125 DCPRINT SAY aArray[i,23]    ///last8
		IF DCPAGEEJECT()
			_PrintPayUpsHdr(@nPage)
		ENDIF
		SKIPALINE()
		IF DCPAGEEJECT()
			_PrintPayUpsHdr(@nPage)
		ENDIF
		@ DC_PRINTERROW()+1,0,DC_PRINTERROW()+1,140 DCPRINT LINE
		IF DCPAGEEJECT()
			_PrintPayUpsHdr(@nPage)
		ENDIF
	NEXT

	PRINTOFF(oPrinter)
	return nil


static function _FALLSINRANGE(dQdate,dEnddate,@dExpdate)
	local dStart,nAdddays
	SELECT SERVCUST
	dStart:=SERVCUST->DELDATE
	nAddDays:=SERVCUST->TERM*30.4
	dExpdate:=dStart+nAdddays
	IF dExpdate >=dQdate .AND. dExpdate <=dEnddate
		RETURN TRUE
	ENDIF
	RETURN FALSE



static function _nextbDay(dDate,dQdate,dEnddate,dNextdate, nAge)
	local cDate,cCurrdate:=dtoc(dQdate),cxDate,cEnddate:=dtoc(dEnddate),dxNextdate


	cDate:=cxDate:=dtoc(dDate)
	cDate:=substr(cDate,1,6)+substr(cCurrdate,7,2)
	cxDate:=substr(cDate,1,6)+substr(cEnddate,7,2)


	dNextdate:=ctod(cDate)
	dxNextdate:=ctod(cxDate)

	IF dNextdate >=dQdate .AND. dNextdate <= dEnddate
		nAge:=(dNextdate-dDate) / 365
		IF nAge < 92
		 RETURN TRUE
		endif
	ENDIF

	IF dxNextdate >=dQdate .AND. dxNextdate <= dEnddate
		nAge:=(dxNextdate-dDate) / 365
		IF nAge < 92
			dNextdate:=dxNextdate              /// stuff new next bday
		 RETURN TRUE
		endif
	ENDIF

	return FALSE



static function _LookCustDealInfo(cLast8)
	LOCAL nRec
	nRec:=SERVRESV->(RECNO())
	SELECT SERVCUST
	ORDSETFOCUS('SERVLST8')
	SERVCUST->(DBSEEK(cLast8))
	IF !FOUND()
		 RETURN FALSE
	ENDIF
	LOOKCUSTDEAL()
	SELECT SERVRESV
	SERVRESV->(DC_DBGOTO(nRec))
	RETURN NIL



FUNCTION _GETSERVCUSTINFO(cLast8,cField)
	local nPos:=0
	SELECT SERVCUST
	ORDSETFOCUS('SERVLST8')
	SERVCUST->(DBSEEK(cLast8))
	IF !FOUND()
		 RETURN ''
	ENDIF
	nPos:=FIELDPOS(cField)
	RETURN FIELDGET(nPos)

  static function _GETSERVCUSTPAYINFO(cLast8)
	local dStart,nAdddays
	SELECT SERVCUST
	ORDSETFOCUS('SERVLST8')
	SERVCUST->(DBSEEK(cLast8))
	IF !FOUND()
		 RETURN ''
	ENDIF
	dStart:=SERVCUST->DELDATE
	nAddDays:=SERVCUST->TERM*30.4
	//WTF SERVCUST->LASTNAME,SERVCUST->DELDATE,nAddDays,SERVCUST->TERM,dStart+nAdddays


	RETURN dStart+nAdddays

 static function _AddSalesNotes(nRec)
	LOCAL GETLIST:={},GETOPTIONS,cComment:='' ,cExist:='',xStatus
	DCGETOPTIONS SAYWIDTH 0 SAYFONT '10.Courier New' GETFONT '10.Courier New'   Autoresize
	SELECT SERVCUST
	SERVCUST->(DBGOTO(nRec))
	cExist:=alltrim(SERVCUST->SALECOMM)
	@ 1,0 DCMULTILINE cExist  SIZE 80,10  MAXLINES 5 EDITPROTECT{||.T.}    ;
  		 NOHORIZSCROLL   ;
       font '10.Arial' ;
		 COLOR 1,BD_LEMONCREME

	//cComment:=cComment+DTOC(DATE())+' '+TIME()+CRLF

	@ 12,0 DCMULTILINE cComment  SIZE 80,15  MAXLINES 14     ;
		 NOHORIZSCROLL ;
		 font '10.Arial' ;
		 COLOR 1,BD_LINEN


	DCREAD GUI TO xStatus MODAL ADDBUTTONS FIT OPTIONS GETOPTIONS TITLE 'Sales Comments Existing and New'
	IF !xStatus
		RETURN NIL
	ENDIF

	cComment:=ALLTRIM(SERVCUST->SALECOMM)+CRLF+DTOC(DATE())+' '+TIME()+USERINIT()+CRLF+cComment+CRLF

	IF SERVCUST->(DC_RECLOCK())
		REPLACE SERVCUST->SALECOMM WITH alltrim(cComment)
		SERVCUST->(DBRUNLOCK(nRec))
	ENDIF
	RETURN NIL

 function getvinsol()
 local getlist:={},getoptions, i , cCell, cPhone, cZip , cSm, cID, xStatus:=.t. , cEmail , cAlt
 local aArray:={},  oDialog   , nCount:=0   , oExcel  ,lOk2add

 aArray:=dc_excel2array((GETFLG('ACCTPATH')+'EXCEL.XLS'))
 IF empty(aArray)
	return nil
 ENDIF



	oDialog:=dc_waiton('Gathering Vin Solutions data from spreadsheet.')
	FOR i:=2  TO len(aArray)
		cPhone:=''
		cCell:=''
		cAlt:=''
		cZip:=''
		cEmail:=''
		IF valtype(aArray[i,7])='N'
			cZip:=alltrim(str(aArray[i,7]))
			IF substr(cZip,5,1)='.'
				cZip:='0'+substr(cZip,1,4)
			ENDIF
		ENDIF
		IF valtype(aArray[i,13])='C'
		 IF !aArray[i,13] $('Dup')
			lOk2Add:=.t.
			SELECT INETUPS
			ORDSETFOCUS('INETUPS')
			IF valtype(aArray[i,10])='C'
			 cEmail:=UPPER(aArray[i,10])
			 cEmail:=alltrim(cEmail)
			 SEEK cEmail
			 IF FOUND()
				lOk2Add:=.f.
			 ENDIF
			ENDIF
			IF Valtype(aArray[i,9]) ='N'
			 IF lOk2Add
			  cPhone:=str(aArray[i,9])
			  cPhone:=alltrim(cPhone)
			  cPhone:=substr(cPhone,1,10)
			  ORDSETFOCUS('INETPHON')
			  seek cPhone
			  IF FOUND()
				 lOk2Add:=.f.
			  ENDIF
			 ENDIF
			 IF lOk2Add
			  ORDSETFOCUS('INETCELL')
			  SEEK cPhone
			  IF FOUND()
				 lOk2Add:=.f.
			  ENDIF
			 ENDIF
			 IF lOk2Add
			  ORDSETFOCUS('INETEVE')
			  SEEK cPhone
			  IF FOUND()
				 lOk2Add:=.f.
			  ENDIF
			 ENDIF
			ENDIF

			IF Valtype(aArray[i,8]) ='N'
			 IF lOk2Add
			  cCell:=str(aArray[i,8])
			  cCell:=alltrim(cCell)
			  cCell:=substr(cCell,1,10)
			  ORDSETFOCUS('INETPHON')
			  SEEK cCell
			  IF FOUND()
				 lOk2Add:=.f.
			  ENDIF
			 ENDIF
			 IF lOk2Add
			  ORDSETFOCUS('INETCELL')
			  SEEK cCell
			  IF FOUND()
				 lOk2Add:=.f.
			  ENDIF
			 ENDIF
			 IF lOk2Add
			  ORDSETFOCUS('INETEVE')
			  SEEK cCell
			  IF FOUND()
				 lOk2Add:=.f.
			  ENDIF
			 ENDIF
			ENDIF
			IF Valtype(aArray[i,17]) ='N'
			 IF lOk2Add
			  cAlt:=str(aArray[i,17])
			  cAlt:=alltrim(cAlt)
			  cAlt:=substr(cAlt,1,10)
			  ORDSETFOCUS('INETPHON')
			  SEEK cAlt
			  IF FOUND()
				 lOk2Add:=.f.
			  ENDIF
			 ENDIF
			 IF lOk2Add
			  ORDSETFOCUS('INETCELL')
			  SEEK cAlt
			  IF FOUND()
				 lOk2Add:=.f.
			  ENDIF
			 ENDIF
			 IF lOk2Add
			  ORDSETFOCUS('INETEVE')
			  SEEK cAlt
			  IF FOUND()
				 lOk2Add:=.f.
			  ENDIF
			 ENDIF
			ENDIF
		endif
		IF EMPTY(cEmail+cPhone+cCell+cAlt)
				lOk2Add:=.f.
		ENDIF
		IF lOk2Add
				cSm:=''
				cId:=''
				SELECT SM
				IF valtype(aArray[i,1])='C'
				 LOCATE FOR SM->INETNAME=ALLTRIM(aArray[i,1])
				 IF found()
					cSm:=SM->SALESMAN
					cId:=SM->ID
				 ENDIF
				endif
				SELECT INETUPS
				IF INETUPS->(DC_ADDREC())
					IF !empty(aArray[i,3])
						 IF valtype(aArray[i,3])='C'
					         REPLACE INETUPS->LASTNAME WITH aArray[i,3]
						 endif
					endif
					IF !empty(aArray[i,2])
						IF valtype(aArray[i,2])='C'
					         REPLACE INETUPS->FIRSTNAME WITH aArray[i,2]
						endif
					endif
					IF !empty(aArray[i,4])
					 IF valtype(aArray[i,4])='C'
					      REPLACE INETUPS->STREET WITH aArray[i,4]
					 endif
					endif


					IF !empty(aArray[i,5])
						  IF valtype(aArray[i,5])='C'
					       REPLACE INETUPS->CITY WITH aArray[i,5]
						  endif
					endif
					IF !empty(aArray[i,6])
						IF valtype(aArray[i,6])='C'
					     REPLACE INETUPS->STATE WITH aArray[i,6]
						endif
					endif
					REPLACE INETUPS->ZIPCODE WITH cZip

					IF !empty(cEmail)
					 REPLACE INETUPS->EMAIL WITH cEmail
					endif

					REPLACE INETUPS->DAYPHONE WITH cPhone
					REPLACE INETUPS->EVEPHON WITH cAlt
					REPLACE INETUPS->CELL WITH cCell

					IF !empty(aArray[i,1])
					  REPLACE INETUPS->CONLAST WITH aArray[i,1]
					endif
					IF !empty(aArray[i,2] )
						IF !EMPTY(aArray[i,3])
							IF valtype(aArray[i,2])='C'
								IF valtype(aArray[i,3])='C'
					            REPLACE INETUPS->FULLNAME WITH ALLTRIM(aArray[i,2])+' '+aArray[i,3]
								endif
							endif
						ENDIF
					endif
					IF !empty(aArray[i,12])
					 REPLACE INETUPS->SOURCE WITH aArray[i,12]
					endif
					IF !empty(aArray[i,11])
					 REPLACE INETUPS->TIMESTAMP WITH dtoc(aArray[i,11])
					endif
					REPLACE INETUPS->Sm WITH cSm
					REPLACE INETUPS->SALESID WITH cId
					REPLACE INETUPS->DATERECVD WITH DATE()
				ENDIF
				nCount++
			ENDIF
		ENDIF
	NEXT
	DC_IMPL(oDialog)
	DC_WINALERT('Number of Leads Added  '+alltrim(str(nCount))+' '+'Press any key to continue.')

return nil

FUNCTION DC_Excel2Array( cExcelFile )

LOCAL oExcel, oSheet, oBook, aValues:={}

#if XPPVER > 1900000
  // Create the "Excel.Application" object
  oExcel := CreateObject("Excel.Application")
  IF Empty( oExcel )
    DC_WinAlert( "Excel is not installed. You need Excel to run this function. Contact your manager." )
    RETURN nil
  ENDIF
#else
  DC_WinAlert('This feature is available in Xbase++ 1.9 and later only!')
  RETURN nil
#endif

oExcel:Visible := .f.

// Load a Workbook from an .XLS file

IF !FExists(cExcelFile)
  //DC_WinAlert( 'File does not exist:' + Chr(13) + cExcelFile )
  RETURN nil
ENDIF

oBook := oExcel:Workbooks:Open(cExcelFile)

oSheet := oBook:activeSheet

aValues := oBook:workSheets(1):usedRange:value

oBook:close()
oBook:destroy()

// Quit Excel
oExcel:Quit()
oExcel:Destroy()

RETURN aValues




static function abomemowrite(cTargetFile,cBuffer)
Local nTarget := FCreate( cTargetFile, FC_NORMAL )
if nTarget == -1
    return .F.
  else
    FWrite( nTarget, cBuffer)
    FClose(ntarget)
endif
return .T.




FUNCTION DAILYUPLOG()
	LOCAL GETLIST:={},GETOPTIONS, dDateIn:=DATE(), xStatus, dDateto:=DATE(), nPage:=0  , cSoldinfo:=''
	local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
	LOCAL aLATTR:=ARRAY(GRA_AL_COUNT) , nXcoord:=0 , clUpmail:='N' , cCell, cWork , cHome
	TOP_MAR(2)
	BOT_MAR(2)
	PAGE_LEN(62)
	aLATTR[GRA_AL_COLOR]:=GRA_CLR_BLACK

	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE

	IF !OPENUPSFILES()
		RETURN NIL
	ENDIF

	IF USE_UDF('UPMAIL',.T.)
	 DBGOTOP()
   ELSE
    CLOSE ALL
   ENDIF

	SELECT NEWUPS
	ORDSETFOCUS('NEWUPDT')
	NEWUPS->(DBGOTOP())

	@ 10,0 DCSAY 'Enter the Date to Run the Up Log From' get dDatein
	@ 11,0 DCSAY 'Enter the Date to Run the Up Log To  ' get dDateto
	@ 12,0 DCSAY 'Create Mail Merge File Y/N ?' get clUpmail picture '@! L'
	DCREAD GUI MODAL TO xStatus ADDBUTTONS ENTEREXIT OPTIONS GETOPTIONS FIT TITLE 'Daily Up Log' EVAL{|o| SETAPPWINDOW(o)}
	IF !xstatus
		RETURN NIL
	ENDIF
	IF clUpmail='Y'
		select upmail
		upmail->(dbzap())
	ENDIF

	IF !Print_Choice('Daily Up Log', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
   ENDIF
   IF !PrintOn('Daily Up Log', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
   ENDIF

	SELECT NEWUPS
	SEEK dDatein
	IF !FOUND()
		DC_WINALERT('No Ups recorded for this date!')
		close all
		RETURN NIL
	ENDIF
	_UPLOGHDR(@nPage,dDatein,dDateto)
	DO WHILE NEWUPS->DATEIN <= dDateto
		nXCOORD:=DC_PRINTERROW()+1.2
		IF DC_PRINTERROW() >= 56
			DCPRINT EJECT
			_UPLOGHDR(@nPage,dDatein,dDateto)
			nXCOORD:=DC_PRINTERROW()+1.2
		ENDIF
		@ nXCOORD,0,nXCOORD+4.3,132 DCPRINT BOX LINEATTR ALATTR LINEWIDTH 1
		@ DC_PRINTERROW()+1.5,0 DCPRINT SAY NEWUPS->SALESMAN
		@ DC_PRINTERROW(),5 DCPRINT SAY ALLTRIM(SUBSTR(NEWUPS->LASTNAME,1,20))+','+ALLTRIM(SUBSTR(NEWUPS->FIRSTNAME,1,10)) FONT '8.Courier New Bold'
		@ DC_PRINTERROW(),27 DCPRINT SAY NEWUPS->APPOINT  PICTURE 'Y'    FONT {|| IIF(NEWUPS->APPOINT ,'10.Courier New Bold' ,'8,Courier New' )}
		@ DC_PRINTERROW(),33 DCPRINT SAY NEWUPS->INTEREST
		IF !EMPTY(NEWUPS->BEBACK1)
			IF NEWUPS->UPTYPE=1
		    @ DC_PRINTERROW(),43 DCPRINT SAY '3'
			ENDIF
			IF NEWUPS->UPTYPE=2
		    @ DC_PRINTERROW(),43 DCPRINT SAY '4'
			ENDIF
		  ELSE
		   @ DC_PRINTERROW(),43 DCPRINT SAY NEWUPS->UPTYPE
		ENDIF
		IF NEWUPS->ORIGSOURCE ='I'
			@ DC_PRINTERROW(),45 DCPRINT SAY NEWUPS->ORIGSOURCE  FONT '12.Courier New'
		 else
		   @ DC_PRINTERROW(),45 DCPRINT SAY NEWUPS->ORIGSOURCE
		ENDIF
		@ DC_PRINTERROW(),55 DCPRINT SAY NEWUPS->ADSOURCE
		@ DC_PRINTERROW(),65 DCPRINT SAY NEWUPS->TOMGR
		@ DC_PRINTERROW(),72 DCPRINT SAY NEWUPS->NEWUSED
		@ DC_PRINTERROW(),78 DCPRINT SAY NEWUPS->DEMOED   PICTURE 'Y'  FONT {|| IIF(NEWUPS->DEMOED ,'10.Courier New Bold' ,'8,Courier New' )}
		@ DC_PRINTERROW(),84 DCPRINT SAY NEWUPS->WRITEUP  PICTURE 'Y'  FONT {|| IIF(NEWUPS->WRITEUP ,'10.Courier New Bold' ,'8,Courier New' )}
		@ DC_PRINTERROW(),89 DCPRINT SAY NEWUPS->TRADEYR+NEWUPS->TRADEMAKE
		IF NEWUPS->UPSTATUS='S'
			cSoldinfo:=''
			_GETSOLDINFO(@cSoldinfo)
			@ DC_PRINTERROW(),100 DCPRINT SAY 'S O L D ' +NEWUPS->STKQUOTED +cSoldinfo FONT '10.Courier New Bold'
		ENDIF
		IF DCPAGEEJECT()
			_UPLOGHDR(@nPage,dDatein,dDateto)
		ENDIF
		@ DC_PRINTERROW()+1,5 DCPRINT SAY NEWUPS->STREET
		@ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->DATEIN
		@ DC_PRINTERROW(),84 DCPRINT SAY NEWUPS->TRADEDESC
		IF DCPAGEEJECT()
			_UPLOGHDR(@nPage,dDatein,dDateto)
		ENDIF
		@ DC_PRINTERROW()+1,5 DCPRINT SAY NEWUPS->CITYSTATE+' '+NEWUPS->ZIPCODE
		@ DC_PRINTERROW(),100 DCPRINT SAY NEWUPS->COMMENTS FONT '6.Courier New Bold'
		IF DCPAGEEJECT()
			_UPLOGHDR(@nPage,dDatein,dDateto)
		ENDIF
		cHome:=NEWUPS->UPID
		_formatphn(@cHome)
		@ DC_PRINTERROW()+1,5 DCPRINT SAY NEWUPS->AREA+'-'+cHome
		cCell:=NEWUPS->CELLPHONE
		cWork:=NEWUPS->workphn
		_formatphn(@cCell)
		_formatphn(@cWork)
		@ DC_PRINTERROW(),25 DCPRINT SAY 'Cell '+ cCell
		@ DC_PRINTERROW(),45 DCPRINT SAY 'Work '+ cWork

		@ DC_PRINTERROW(),80 DCPRINT SAY NEWUPS->EMAIL
		IF DCPAGEEJECT()
			_UPLOGHDR(@nPage,dDatein,dDateto)
		ENDIF
		IF clUpmail='Y'
			makeupmail()      // add record to mailfile
		ENDIF
		SKIP ALIAS NEWUPS
		IF NEWUPS->(EOF())
			EXIT
		ENDIF
	ENDDO
	PRINTOFF(oPrinter)
	CLOSE ALL
	IF clUpmail='Y'
		EditUpMail()
	ENDIF
	close all
	RETURN NIL

	function _formatphn(cPhn)
	  cPhn:=alltrim(cPhn)
	  IF len(cPhn)=7
		cPhn:=substr(cPhn,1,3)+'-'+substr(cPhn,4,4)
		return cPhn
	  ENDIF
	  IF len(cPhn)=10
		cPhn:=substr(cPhn,1,3)+'-'+substr(cPhn,4,3)+'-'+SUBSTR(cPhn,7,4)
		return cPhn
	  ENDIF
	 return cPhn

	STATIC FUNCTION _GETSOLDINFO(cSoldinfo)
		SELECT NCI
		SEEK NEWUPS->STKQUOTED
		IF FOUND()
			cSoldinfo:=NCI->YEAR+' '+ALLTRIM(NCI->MAKE)+' '+ALLTRIM(NCI->MODEL)
			RETURN cSoldinfo
		ENDIF
		SELECT UCL01
		SEEK NEWUPS->STKQUOTED
		IF FOUND()
			cSoldinfo:=UCL01->YEAR+' '+ALLTRIM(UCL01->MAKE)+' '+ALLTRIM(UCL01->MODEL)
			RETURN cSoldinfo
		ENDIF
		RETURN NIL


	STATIC FUNCTION _UPLOGHDR(nPage,dDatein,dDateto)
		nPage++
		@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Daily Up Log For '+ dtoc(dDateIn) +' to ' +dtoc(dDateTo)+'  Page '+ alltrim(str(nPage))
		@ DC_PRINTERROW(),90 DCPRINT SAY CDOW(dDateIn) FONT '20.Courier New Bold'
		@ DC_PRINTERROW(),110 DCPRINT SAY CDOW(dDateTo) FONT '20.Courier New Bold'
		@ DC_PRINTERROW()+2,0 DCPRINT SAY 'SM'
		@ DC_PRINTERROW(),5  DCPRINT SAY 'Name/Address/Phone'
		@ DC_PRINTERROW(),27 DCPRINT SAY 'Appt'
		@ DC_PRINTERROW(),33 DCPRINT SAY 'Intrst'
		@ DC_PRINTERROW(),41 DCPRINT SAY 'Type'
		@ DC_PRINTERROW(),45 DCPRINT SAY 'OrgSrc'
		@ DC_PRINTERROW(),55 DCPRINT SAY 'AdvSrc'
		@ DC_PRINTERROW(),65 DCPRINT SAY 'ToMgr'
		@ DC_PRINTERROW(),72 DCPRINT SAY 'N/U'
		@ DC_PRINTERROW(),78 DCPRINT SAY 'Demo'
		@ DC_PRINTERROW(),84 DCPRINT SAY 'WtUp'
		@ DC_PRINTERROW(),88 DCPRINT SAY 'Trade'
		@ DC_PRINTERROW(),100 DCPRINT SAY 'Comments'
		RETURN NIL





/* This function is designed to do a wholesale change in the service customer database from one salespersons initials to another
	The original selling saleperson is added to the comment field */

Function reassignsales()
local getlist:={}
local getoptions , xStatus  , cProgram , oDialog
local cOldSm:=space(3), cNewSm:=space(3)
local clCon:='Y'
local clGoOn:='Y'
local cNewName, nCount:=0
DCGEToptions saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize

cProgram:="REASSIGN"
IF .NOT.SECURITY(cProgram)
	dc_winalert( 'SECURITY VIOLATION'   )
	RETURN NIL
ENDIF


/// OPEN FILEs

IF USEDB({'SERVCUST','SM'})
	SERVCUST->(DBGOTOP())
ELSE
	DC_WINALERT('Cannot open files')
	CLOSE ALL
	RETURN NIL
ENDIF


odialog:=dc_waiton('Now Backing up Service Customer File ')
copy file (GETFLG('ACCTPATH')+'SERVCUST.dbf') to (GETFLG('ACCTPATH')+'SERVCUST.bak')
dc_impl(oDialog)
do while clCon='Y'
	@ 5,0 DCSAY 'This Program is designed to do a wholesale change of the Saleperson ID in the Service Customer file.'
	@ 6,0 DCSAY 'The Sales Person ID Entered as New will overlay the SalePerson ID currently stored.'
	@ 10,0 DCSAY 'Enter the NEW Sales person ID to Use.       ' get cNewSM valid{|| _chksm(cNewSM)}
	@ 11,0 DCSAY 'Enter the Old SalesPerson ID to be replaced.' get cOldSM
	dcread gui modal enterexit to xstatus fit options getoptions Title 'Salesperson ID Change'   eval{|o| setappwindow(o)}
	IF !xStatus
		return nil
	ENDIF
	SELECT SM
	SEEK cNewSM
	cNewName:=alltrim(SM->Name)
	@ 10,0 DCSAY 'Verify Please..  Replace '+cOldSM+' With '+cNewSM +' '+cNewName
	@ 11,0 DCSAY 'Enter Y to verify or N to exit    ' get clGoOn picture '@! L'
	dcread gui modal enterexit to xstatus fit options getoptions Title 'Salesperson ID Change'   eval{|o| setappwindow(o)}
	IF !xStatus
		return nil
	ENDIF
	IF clGoOn='Y'
		_changeSM(cOldSM,cNewSM,cNewName,@nCount)
	ENDIF
	@ 5,0 DCSAY 'Records Changed from '+cOldSM +' to '+ cNewSM +' '+alltrim(str(nCount))
	@ 10,0 DCSAY 'Change Another Y/N' get clCon picture '@! L'
	dcread gui modal enterexit to xstatus fit options getoptions Title 'Salesperson ID Change'   eval{|o| setappwindow(o)}
	IF !xStatus
		return nil
	ENDIF
enddo
return nil

STATIC FUNCTION _chksm(cNewSM)
	SELECT SM
	LOCATE FOR SM->ID=cNewSM
	IF found()
		return TRUE
	ENDIF
	dc_winalert('SM Id is invalid')
	return FALSE

static function _changeSM(cOldSM,cNewSM,cNewname,nCount)
local getlist:={},oDialog
select SERVCUST
SERVCUST->(DBGOTOP())

odialog:=dc_waiton(cOldSm+' to '+cNewSM)
DO WHILE !SERVCUST->(EOF())
	  IF SERVCUST->SM=cOldSM
	  	IF SERVCUST->(dc_RECLOCK())
	  		REPLACE SERVCUST->SM WITH cNewSM
	  		REPLACE SERVCUST->SMNAME WITH cNewname+' OldSP '+SERVCUST->SMNAME
	  		UNLOCK
	  		nCount++
	    ENDIF
	  ENDIF
	  skip alias SERVCUST
ENDDO
DC_IMPL(oDialog)
RETURN nCount


/* this function is designed to list customers whose lease or retail contract is about to expire.
	 It uses a from-to date choice to isolate approx dates that the contracts will expire.
	 The user can segregate leases from retail and can input the number of days to expiration.  */

function premarketlook()
	local getlist:={}
	local nDaystoexp:=0 , cSM:='',cCurName:=''
	local cContype:=' '
	local dExdate:=ctod('  /  /  ')
	local getoptions , oDialog , xStatus , nFoundCount:=0, nTotalFoundCount:=0 , nTotRetOps:=0,nTotLseOps:=0
	LOCAL nOrient:=2, nDefmode:=4,nCopies:=1, cOutfile:='',cSentfont:='8.Courier New', nPage:=0 , oPrinter

	//local dStart:=date()
	TOP_MAR(3)
	BOT_MAR(2)
	PAGE_LEN(62)
	nPAGE:=0
	DCGEToptions saywidth 0 sayfont '12.Courier New' getfont '12.Courier New' autoresize

	IF UseDb({'SERVCUST','EMPMAST'})
		SERVCUST->(ORDSETFOCUS('SERVSM'))
	ELSE
		CLOSE ALL
		RETURN NIL
	ENDIF

	/*
	if use_udf(GETFLG('HISTPATH')+'SERVCUST',.f.)
		goto top
		oDialog:=dc_waiton('Now Sorting records by Sales Person')
		index on SERVCUST->sm to (GETFLG('HISTPATH')+'servsm')
		set index to servsm
		dc_impl(oDialog)
	else
		close all
		return nil
	endif  */

	@ 10,0 DCSAY 'Enter L for Leases or R for Retails OR B for Both' get cContype valid{|| iif(cContype $('LRB'),.t.,.f.)}  PICTURE '@!'

	@ 12,0 DCSAY 'Enter days to contract end. EG 120 would be to contact 4 months out. ' get nDaystoexp
	//@ 13,0 DCSAY 'Enter the start date. ' get dStart
	dcread gui to xstatus modal enterexit fit options getoptions title 'Pre term Contact List';
		eval{|o| setappwindow(o)}

	if !xstatus
		close all
		return nil
	endif

	IF cContype='B'
		cContype:='LR'
	ENDIF

	IF !Print_Choice( 'Lease/Retail Expiration', @nCopies,' ',.F.,@nOrient,@cSentfont,@nDefMode,cOutfile )
			 RETURN NIL
	ENDIF

	IF nOrient=1
			PAGE_LEN(80)
	ENDIF

	IF nDefMode=8
		PAGE_LEN(32000)
	ENDIF


	IF !PrintOn('Lease/Retail Expiration', oPrinter, nOrient, cSentfont, nCopies)
		  RETURN NIL
	ENDIF
	premarkhdr(@nPage)

	DC_HOURGLASSON()
	select SERVCUST
	do while !SERVCUST->(eof())
		cSM:=SERVCUST->SM
		cCurName:=''
		SELECT EMPMAST
		EMPMAST->(DBSEEK(cSM))
		IF found()
			cCurName:=SUBSTR(EMPMAST->LASTNAME,1,9)+' '+SUBSTR(EMPMAST->FIRST,1,4)
		ENDIF
		SELECT SERVCUST
		nFoundcount:=0

		DO WHILE SERVCUST->SM=cSM
		 if SERVCUST->soldtype $(cContype)
			if !empty(SERVCUST->deldate)
				dExdate:=SERVCUST->deldate+(30*SERVCUST->term)
				if dExdate - date() > 0
					if dExdate - date() < nDaystoexp
						nFoundcount++
						nTotalFoundCount++
						IF SERVCUST->SOLDTYPE='L'
							nTotLseOps++
						ENDIF
						IF SERVCUST->SOLDTYPE='R'
							nTotRetOps++
						ENDIF
						prnttherec(@nPAGE,cCurName)
					endif
				endif
			endif
		 endif
		 skip alias SERVCUST
		ENDDO
		IF nFoundCount > 0
		 skipaline()
		 if dcpageeject()
		  premarkhdr(@nPAGE)
	    endif
		 @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Opportunities Found ' + alltrim(str(nFoundCount))

		 DCPRINT EJECT
		 premarkhdr(@nPage)
		endif
	enddo
	DCPRINT EJECT

	@ DC_PRINTERROW()+1,0 DCPRINT SAY '       Total Opportunities Found '
	@ DC_PRINTERROW(),50 DCPRINT SAY nTotalFoundCount PICTURE '99999'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY ' Total Lease Opportunities Found '
	@ DC_PRINTERROW(),50 DCPRINT SAY nTotLseOps PICTURE '99999'

	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Retail Opportunities Found '
	@ DC_PRINTERROW(),50 DCPRINT SAY nTotRetOps PICTURE '99999'


	DC_HOURGLASSOFF()
	PRINTOFF(oPrinter)
	close all
return nil

static function prnttherec(nPAGE,cCurName)
	@ dc_printerrow()+1,0 DCPRINT SAY ' '
	if dcpageeject()
		premarkhdr(@nPAGE)
	endif


	@ DC_PRINTERROW()+1,0 DCPRINT SAY SERVCUST->SM   FONT '8.Courier New Bold'
	@ DC_PRINTERROW(),15 DCPRINT SAY substr(SERVCUST->lastname,1,15) FONT '8.Courier New Bold'
	@ DC_PRINTERROW(),35 DCPRINT SAY  SERVCUST->STREET
	@ DC_PRINTERROW(),65 DCPRINT SAY  SERVCUST->YEAR
	@ DC_PRINTERROW(),82 DCPRINT SAY  SERVCUST->DELDATE
	@ DC_PRINTERROW(),95 DCPRINT SAY  SERVCUST->SALETYPE
	@ DC_PRINTERROW(),105 DCPRINT SAY  SERVCUST->hphonealph picture '@R 999.999.9999' FONT '8.Courier New Bold'
	if dcpageeject()
		premarkhdr(@nPAGE)
	endif




	@ DC_PRINTERROW()+1,0 DCPRINT SAY cCurName
	@ DC_PRINTERROW(),15 DCPRINT SAY substr(SERVCUST->firstname,1,10)
	@ DC_PRINTERROW(),35 DCPRINT SAY  ALLTRIM(SERVCUST->CITYSTATE)+SERVCUST->STATE+' '+SERVCUST->ZIP
	@ DC_PRINTERROW(),65 DCPRINT SAY SERVCUST->MAKE
	@ DC_PRINTERROW(),82 DCPRINT SAY SERVCUST->deldate+(SERVCUST->term*30) FONT '8.Courier New Bold'
	@ DC_PRINTERROW(),95 DCPRINT SAY  SERVCUST->SOLDTYPE
	@ DC_PRINTERROW(),105 DCPRINT SAY  SERVCUST->wphonealph picture '@R 999.999.9999' FONT '8.Courier New Bold'
	if dcpageeject()
		premarkhdr(@nPAGE)
	endif


	@ DC_PRINTERROW()+1,0 DCPRINT SAY SUBSTR(SERVCUST->SMNAME,1,15)
	@ DC_PRINTERROW(),15 DCPRINT SAY SERVCUST->VIN
	@ DC_PRINTERROW(),35 DCPRINT SAY ALLTRIM(SERVCUST->EMAIL)
	@ DC_PRINTERROW(),65 DCPRINT SAY SERVCUST->CARDESC
	@ DC_PRINTERROW(),82 DCPRINT SAY SERVCUST->LASTSERV
	@ DC_PRINTERROW(),95 DCPRINT SAY  SERVCUST->TERM
	@ DC_PRINTERROW(),105 DCPRINT SAY  SERVCUST->cell picture '@R 999.999.9999' FONT '8.Courier New Bold'
	if dcpageeject()
		premarkhdr(@nPAGE)
	endif

return nil



static function premarkhdr(nPAGE)
	npage:=npage++
	@ dc_printerrow()+1,0 DCPRINT SAY 'Retails/Leases soon to expire.  Page '+ alltrim(str(nPage))

	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Sales Id' //SCUST->SM   FONT '8.Courier New Bold'
	@ DC_PRINTERROW(),15 DCPRINT SAY 'LastName' //substr(SERVCUST->lastname,1,15) FONT '8.Courier New Bold'
	@ DC_PRINTERROW(),35 DCPRINT SAY  'Street' //SERVCUST->STREET
	@ DC_PRINTERROW(),65 DCPRINT SAY  'Year' // SERVCUST->YEAR
	@ DC_PRINTERROW(),82 DCPRINT SAY  'DelivDate' //SERVCUST->DELDATE
	@ DC_PRINTERROW(),95 DCPRINT SAY  'New Used' //SERVCUST->SOLDTYPE
	@ DC_PRINTERROW(),105 DCPRINT SAY 'HomePhn' //SERVCUST->hphonealph picture '@R 999.999.9999' FONT '8.Courier New Bold'




	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'SpName' //rcCurName
	@ DC_PRINTERROW(),15 DCPRINT SAY 'FirstName' //substr(SERVCUST->firstname,1,10)
	@ DC_PRINTERROW(),35 DCPRINT SAY 'City State Zip' //ALLTRIM(SERVCUST->CITYSTATE)+SERVCUST->STATE+' '+SERVCUST->ZIP
	@ DC_PRINTERROW(),65 DCPRINT SAY 'Make' //SERVCUST->MAKE
	@ DC_PRINTERROW(),82 DCPRINT SAY 'EXPIRE DATE' FONT '8.Courier New Bold'
	@ DC_PRINTERROW(),95 DCPRINT SAY  'FinType' //SERVCUST->SALETYPE
	@ DC_PRINTERROW(),105 DCPRINT SAY 'WorkPhn' //SERVCUST->wphonealph picture '@R 999.999.9999' FONT '8.Courier New Bold'


	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'SoldBy' //SERVCUST->SMNAME
	@ DC_PRINTERROW(),15 DCPRINT SAY 'Vin' //SERVCUST->VIN
	@ DC_PRINTERROW(),35 DCPRINT SAY 'Email' // ALLTRIM(SERVCUST->EMAIL)
	@ DC_PRINTERROW(),65 DCPRINT SAY 'Model' //SERVCUST->CARDESC
	@ DC_PRINTERROW(),82 DCPRINT SAY 'LastServ'//ERVCUST->LASTSERV
	@ DC_PRINTERROW(),95 DCPRINT SAY  'FinTerm' //SERVCUST->TERM
	@ DC_PRINTERROW(),105 DCPRINT SAY  'CellPhn' //SERVCUST->cell picture '@R 999.999.9999' FONT '8.Courier New Bold'

RETURN NIL

/// service customer browser and memo entry
FUNCTION Custnotes()
	LOCAL GetList := {}
	local aPres, oToolBar
	local oBrowse
	LOCAL GETOPTIONS,XSTATUS,cFindme
	IF USEDb({'SERVCUST'})
		ordsetfocus('SERVNMIN')
	ELSE
		DC_WINALERT('Cannot Open Customer File.')
		RETURN NIL
	ENDIF

	DCGETOPTIONS SAYFONT '10.Courier New'  getfont '10.Courier New' ;
		SAYWIDTH 0 ;
		AUTORESIZE CASCADE
	cFindme:=SPACE(10)
	@ 10,20 DCSAY 'Enter Lastname to browse for.' GET cFindme PICTURE '@!' SAYCOLOR 3,0 GETCOLOR 0,3 GETFONT '12.COURIER'
	DCREAD GUI APPWINDOW MAINWINDOW():DRAWINGAREA ENTEREXIT FIT TITLE 'BROWSE SERVICE CUSTOMERS' OPTIONS GETOPTIONS
	SET SOFTSEEK ON
	SELECT SERVCUST

	SEEK cFindme
	SET SOFTSEEK OFF

	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

	@ 1,1 DCTOOLBAR oToolBar                               ;
		SIZE 115, 1.5


	DCADDBUTTON CAPTION 'Edit Sales Comments'                           ;
		SIZE 30                                              ;
		ACTION {|| servcuscomm(), DC_GetRefresh(GetList)}     ;
		PARENT oToolBar                                     ;
		TOOLTIP 'Enter Sales Follow up comments'

	DCADDBUTTON CAPTION 'Enter New Lastname'                           ;
		SIZE 30                                              ;
		ACTION {|| resetname() , DC_GetRefresh(GetList),oBrowse:forcestable()}     ;
		PARENT oToolBar                                     ;
		TOOLTIP 'Enter Sales Follow up cpmments'

	DCADDBUTTON CAPTION 'Stuff Mouse Clipboard with Vin for Paste'                           ;
		SIZE 40                                              ;
		ACTION {|| ystuffvin(ALLTRIM(SERVCUST->VIN)), DC_GetRefresh(GetList)}     ;
		PARENT oToolBar                                     ;
		TOOLTIP 'Stuff Vin for paste'

	DCADDBUTTON CAPTION 'Send eMail '                              ;
		SIZE 15                                              ;
		ACCELKEY {ASC("8")};
		ACTION {||followemail()};
		PARENT oToolBar


/* ----- Create browse ----- */

	@ 3,1 DCBROWSE oBrowse ALIAS 'SERVCUST'                 ;
		SIZE 110,20                                     ;
		PRESENTATION aPres;
		FREEZELEFT {1,2};
		TITLE 'SERVICE CUSTOMER BROWSE AND MAINTAIN';
		EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITEXIT

	DCBROWSECOL FIELD SERVCUST->SMNAME ;
		HEADER "SALESMAN" PARENT oBrowse ;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD SERVCUST->LASTNAME                     ;
		width 15;
		HEADER "LASTNAME"  HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse                ;
		VALID {|c|DC_ReadEmpty(c)}
	DCBROWSECOL FIELD SERVCUST->FIRSTNAME ;
		HEADER "FIRSTNAME" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD SERVCUST->CARDESC ;
		HEADER "DESCRIPTION" PARENT oBrowse ;
		protect{|| .t.}

	DCBROWSECOL FIELD SERVCUST->EMAIL ;
		HEADER "EMAIL" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN  PARENT oBrowse;
		width 25

	DCBROWSECOL FIELD SERVCUST->STREET ;
		HEADER "STREET" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse;
		width 15

	DCBROWSECOL FIELD SERVCUST->CITYSTATE ;
		HEADER "CITYSTATE" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse;
		width 15;

		DCBROWSECOL FIELD SERVCUST->ZIP ;
		HEADER "ZIP" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse

	DCBROWSECOL FIELD SERVCUST->PHONE ;
		HEADER "PHONE" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse;
		width 8

	DCBROWSECOL FIELD SERVCUST->WPHONE ;
		HEADER "WORKPHONE" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse;
		width 8

	DCBROWSECOL FIELD SERVCUST->CELL ;
		HEADER "CELLPHONE" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse ;
		width 8


	DCBROWSECOL FIELD SERVCUST->VIN ;
		HEADER "VIN" PARENT oBrowse   ;
		PROTECT {|| .T.} ;
		width 14

	DCBROWSECOL FIELD SERVCUST->CUSTNO ;
		HEADER "AR/Cust#" PARENT oBrowse ;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD SERVCUST->YEAR ;
		HEADER "YEAR" PARENT oBrowse ;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD SERVCUST->DELDATE ;
		HEADER "DELIVERED" PARENT oBrowse ;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD SERVCUST->PURCHAT ;
		HEADER "PURCHD AT" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse

	DCBROWSECOL FIELD SERVCUST->LASTSERV ;
		HEADER "LASTSERVD" PARENT oBrowse;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD SERVCUST->COLOR ;
		HEADER "COLOR " PARENT oBrowse ;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD SERVCUST->SERCOMP ;
		HEADER "SC COMPANY" PARENT oBrowse ;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD SERVCUST->SERCONT ;
		HEADER "SC TERM" PARENT oBrowse;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD SERVCUST->scterm ;
		HEADER "TERMMOS" PARENT oBrowse;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD SERVCUST->scmiles ;
		HEADER "TERMMILES" PARENT oBrowse ;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD SERVCUST->INSERVMILE ;
		HEADER "INSERVMILES" PARENT oBrowse;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD SERVCUST->STOCKNO ;
		HEADER "STOCK#" PARENT oBrowse   ;
		PROTECT {|| .T.}


	DCBROWSECOL FIELD SERVCUST->MASSMAIL ;
		HEADER "MASSMAIL" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCREAD GUI FIT  BUTTONS DCGUI_BUTTON_EXIT ;
		OPTIONS GETOPTIONS TITLE 'SERVICE CUST BROWSER';
		APPWINDOW MAINWINDOW():DRAWINGAREA
	CLOSE ALL
ReTURN nil
*** END OF EXAMPLE ***

static function resetname()
	local getlist:={}
	local getoptions
	local cFindme:=space(20)
	DCGETOPTIONS SAYFONT '10.Courier New'  getfont '10.Courier New' ;
		SAYWIDTH 0 ;
		AUTORESIZE CASCADE

	@ 5,0 DCSAY 'New name Search'
	@ 10,0 DCSAY 'Enter Lastname to browse for' GET cFindme PICTURE '@!' SAYCOLOR 3,0 GETCOLOR 0,3 GETFONT '12.COURIER'
	DCREAD GUI APPWINDOW MAINWINDOW():DRAWINGAREA ENTEREXIT FIT TITLE 'Name Search' OPTIONS GETOPTIONS
	SET SOFTSEEK ON
	SELECT SERVCUST
	SEEK cFindme
	SET SOFTSEEK OFF
return nil




static function servcuscomm()
	if SERVCUST->(DC_REClock(recno()))
		REPLACE SERVCUST->salecomm WITH DC_MEMOEDIT(SERVCUST->salecomm,2,0,20,79,.T.,,50,,,,,,,)
		SERVCUST->(dbcommit())
		SERVCUST->(DBRUNLOCK())
		return nil
	endif
return nil





//// function to browse stock
FUNCTION SUBROWSTK(cDemodStkNo)
	LOCAL GETLIST:={}, aPres, oBrowse, oToolBar,cIndexOrder:='Model Description Order ',cSearch:=space(20)
	LOCAL ALLOW:='N'  , cProgram , getoptions  , cAvailcode ,oDialog
	LOCAL XSTATUS:=.T. ,cFindme:=SPACE(8)  , nRec
	default cDemodStkNo:=space(8)
	nREC:=0
	cProgram:='UBROWSTK'
	DCGETOPTIONS SAYFONT '12.Courier New' getfont '12.Courier New';
		SAYWIDTH 0 ;
		AUTORESIZE
	cAvailcode:='Y'
	SELECT UCL01

	UCL01->(ordsetfocus('INVIDX03'))
	SET FILTER TO ucl01->AVAILCODE='A'

	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */


@ 1,1 DCSAY 'Current Index Order' get cIndexOrder EDITPROTECT{||.T.}
@ 1,70 DCSAY 'Search For Key' GET cSearch SAYRIGHT PICT '@!' ;
      KEYBLOCK { |a,b,o| DBSELECTAREA('UCL01'),DC_BrowseAutoSeek(a,o,oBrowse)} ;
		GETID 'Srchkey'

@ 2,1 DCSAY 'You may RIGHT CLICK on Column Headers that Begin with ** to change Search Order to that Column '
/* ----- Create browse ----- */

	@ 3.5,1 DCBROWSE oBrowse ALIAS 'UCL01'                 ;
		SIZE 110,20                                     ;
		PRESENTATION aPres  ;
		FREEZELEFT {1,2,3,4} ;
		SORTSCOLOR 0,2;
		SORTUCOLOR 1,6  ;
		DATALINK{|| cDemodStkNo:=UCL01->STOCKNO,DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}


	DCBROWSECOL FIELD UCL01->STOCKNO                     ;
		HEADER "**STOCK#  " HCOLOR 9,3 PARENT oBrowse                  ;
		SORT{|| oDialog:=DC_WAITON('Now Sorting by Stock Number'),cIndexOrder:='Stock Number Order          ',DBSELECTAREA('UCL01'),;
		ucl01->(ordsetfocus('INVIDX01')),;
		DC_IMPL(oDialog),DC_DbGoTop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)}

	DCBROWSECOL FIELD UCL01->LAST6 ;
		HEADER "**Last 6 " HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKRED PARENT oBrowse ;
		WIDTH 6 ;
		SORT{|| oDialog:=DC_WAITON('Now Sorting by Last 6'),cIndexOrder:='Last 6 Of Vin           ',DBSELECTAREA('UCL01'),;
		ucl01->(ordsetfocus('INVLST6')),;
		DC_IMPL(oDialog),DC_DbGoTop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)}

	DCBROWSECOL FIELD UCL01->AVAILCODE ;
		HEADER "AVAILCODE" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse

	DCBROWSECOL FIELD UCL01->YEAR ;
		HEADER "YEAR" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse

   DCBROWSECOL FIELD UCL01->MAKE ;
		HEADER "**MAKE" WIDTH 10 HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse ;
		SORT{|| oDialog:=DC_WAITON('Now Sorting by Make'),cIndexOrder:='Make               ',DBSELECTAREA('UCL01'),;
		ucl01->(ordsetfocus('INVMAKE')),;
		dc_impl(oDialog),;
		DC_DbGoTop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)}

	DCBROWSECOL FIELD UCL01->MODEL ;
		HEADER "**MODEL" WIDTH 10 HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse  ;
		SORT{|| oDialog:=DC_WAITON('Now Sorting by Description'),cIndexOrder:='Model Description',DBSELECTAREA('UCL01'),;
		ucl01->(ordsetfocus('INVIDX03')),;
		dc_impl(oDialog),;
		DC_DbGoTop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)}


	DCBROWSECOL FIELD UCL01->TRIM ;
		HEADER "SERIES" WIDTH 9 HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->COLOR ;
		HEADER "COLOR" WIDTH 6 HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	if ucl01->(isfieldvar('INTCOLOR'))
		DCBROWSECOL FIELD UCL01->INTCOLOR ;
			HEADER "TRIMCOLOR" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	endif
	if ucl01->(isfieldvar('TITLE'))
		DCBROWSECOL FIELD UCL01->TITLE ;
			HEADER "TITLE" HCOLOR 1,5 PARENT oBrowse    COLOR {|| _TITLECOLOR()}
	endif
	DCBROWSECOL FIELD UCL01->VIN ;
		HEADER "VIN" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKRED PARENT oBrowse ;
		WIDTH 12 ;
		PROTECT {|| .T.}
	DCBROWSECOL FIELD UCL01->MILEAGE ;
		HEADER "MILES" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->ENGINE ;
		HEADER "CYL" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->TRANSMISS ;
		HEADER "TRANS" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->PS ;
		HEADER "PS" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->PB ;
		HEADER "PB" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->AIR ;
		HEADER "AC" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->ABS ;
		HEADER "ABS" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->MEDIA ;
		HEADER "MEDIA" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->COST ;
		HEADER "BASECOST" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse ;
		WIDTH 6
	DCBROWSECOL FIELD UCL01->RETAIL ;
		HEADER "LIST PRICE" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse ;
		WIDTH 6
	if ucl01->(isfieldvar('ePRICE'))
		DCBROWSECOL FIELD UCL01->ePRICE ;
			HEADER "SALE Price" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse ;
			WIDTH 6
	endif
	DCBROWSECOL FIELD UCL01->CERTIFIED ;         /// USE THIS FIELD FOR CERTIFIED
		HEADER "CERTIFIED" HCOLOR 1,5 PARENT oBrowse

	DCBROWSECOL FIELD UCL01->COMMENTS ;
		HEADER "COMMENTS" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->VEHTYPE ;
		HEADER "VEHTYPE " HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->SELLPRICE ;
		HEADER "SOLDFOR" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->SLD_DATE ;
		HEADER "DATESOLD" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->SALESMAN ;
		HEADER "SOLDBY" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->CUSTOMER ;
		HEADER "SOLDTO" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->STKDATE ;
		HEADER "DATEIN" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->SOURCE;
		HEADER "SOURCE" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse
	DCBROWSECOL FIELD UCL01->REC_NAME ;
		HEADER "FROM" HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse

	DCREAD GUI ;
		FIT ;
		OPTIONS GETOPTIONS ;
		TITLE 'Used Vehicle Browser' ;
		BUTTONS DCGUI_BUTTON_EXIT

	SET FILTER TO
ReTURN nil

 STATIC FUNCTION _TITLECOLOR()
	IF UCL01->TITLE='Y'
		RETURN {GRA_CLR_BLACK,GRA_CLR_GREEN}
	ENDIF
	IF UCL01->TITLE='L'
		RETURN {GRA_CLR_BLACK,GRA_CLR_YELLOW}
	ENDIF
	IF UCL01->TITLE='N'
		RETURN {GRA_CLR_BLACK,GRA_CLR_RED}
	ENDIF
	IF UCL01->TITLE=' '
		RETURN {GRA_CLR_BLACK,GRA_CLR_WHITE}
	ENDIF
RETURN {GRA_CLR_BLACK,GRA_CLR_WHITE}


FUNCTION BROWAPPRSL()
	LOCAL GETLIST:={}, aPres,oBrowse, oToolBar, cProgram ,oDeskRecord
	LOCAL XSTATUS:=.T. ,cFindme:=SPACE(17) , nRec ,cSearch:='Last Name Search    '
	LOCAL GETOPTIONS
	DCGETOPTIONS SAYWIDTH 0 SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
	nREC:=0
	cProgram:='BROWAPPRSL'
	DCGETOPTIONS SAYFONT '10.Courier New' getfont '10.Courier New';
		SAYWIDTH 0 ;
		AUTORESIZE TABSTOP

  IF !OPENUPSFILES()
	CLOSE ALL
   RETURN NIL
  ENDIF

  SELECT UCAPPRSL
  ORDSETFOCUS('UCAPPLNM')
  UCAPPRSL->(DBGOTOP())
  @ 1,1 DCSAY 'Browse Order  ' get cSearch EDITPROTECT{|| TRUE }
  @ 1,50 DCSAY 'Enter Search Key for Browse Order' GET cFindme SAYRIGHT PICT '@!' ;
      KEYBLOCK {|a,b,o|DC_BrowseAutoSeek(a,o,oBrowse)}



	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

	@ 3,1 DCTOOLBAR oToolBar                               ;
		SIZE 77, 1.5

	DCADDBUTTON CAPTION 'Review Appraisal '                              ;
	SIZE 15                                              ;
	ACTION {|| _FindUpRec(UCAPPRSL->UPID,@oDeskRecord),_appraiseform(NEWUPS->(RECNO()),UCAPPRSL->UPID,UCAPPRSL->SM,@oDeskRecord,.t.),oBrowse:REFRESHALL()};
	PARENT oToolBar


/* ----- Create browse ----- */

	@ 5,1 DCBROWSE oBrowse ALIAS 'UCAPPRSL'                 ;
		SIZE 130,20                                     ;
		PRESENTATION aPres  ;
		FREEZELEFT {1,2} ;
		SORTSCOLOR 0,2 ;
		SORTUCOLOR 1,6   ;
		EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITACROSS

	DCBROWSECOL FIELD UCAPPRSL->APPRDATE                     ;
		HEADER "**DATE" HCOLOR 1,3 PARENT oBrowse   ;
		SORT{|| cSearch:='Appraisal Date     ',UCAPPRSL->(ORDSETFOCUS('UCAPPDAT')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)};
	   EDITPROTECT{|| TRUE } WIDTH 6

	DCBROWSECOL FIELD UCAPPRSL->LASTNAME                     ;
		HEADER "**LASTNAME" HCOLOR 1,5 PARENT oBrowse    ;
		SORT{|| cSearch:='Appraisal Last Name     ',UCAPPRSL->(ORDSETFOCUS('UCAPPLNM')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)} WIDTH 10

	DCBROWSECOL FIELD UCAPPRSL->FIRSTNAME                     ;
		HEADER "FIRSTNAME" HCOLOR 1,5 PARENT oBrowse

	DCBROWSECOL FIELD UCAPPRSL->UPID                     ;
		HEADER "**UP ID" HCOLOR 1,3 PARENT oBrowse   ;
		SORT{|| cSearch:='Appraisal UP ID     ',UCAPPRSL->(ORDSETFOCUS('UCAPPRSL')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)}  ;
	   EDITPROTECT{|| TRUE }

	DCBROWSECOL FIELD UCAPPRSL->YEAR                     ;
		HEADER "YEAR" HCOLOR 1,5 PARENT oBrowse

	DCBROWSECOL FIELD UCAPPRSL->MAKE                     ;
		HEADER "**MAKE" HCOLOR 1,5 PARENT oBrowse  ;
		SORT{|| cSearch:='Appraisal Make     ',UCAPPRSL->(ORDSETFOCUS('UCAPPMAK')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)}

	DCBROWSECOL FIELD UCAPPRSL->MODEL                     ;
		HEADER "**MODEL" HCOLOR 1,5 PARENT oBrowse;
		SORT{|| cSearch:='Appraisal Model     ',UCAPPRSL->(ORDSETFOCUS('UCAPPMOD')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)}

	DCBROWSECOL FIELD UCAPPRSL->SERIES                     ;
		HEADER "SERIES" HCOLOR 1,5 PARENT oBrowse

	DCBROWSECOL FIELD UCAPPRSL->COLOR                     ;
		HEADER "COLOR" HCOLOR 1,5 PARENT oBrowse

	DCBROWSECOL FIELD UCAPPRSL->APPRAISER                     ;
		HEADER "APPRAISER" HCOLOR 1,3 PARENT oBrowse   ;
	   EDITPROTECT{|| TRUE }

	DCBROWSECOL FIELD UCAPPRSL->ACV                     ;
		HEADER "ACV" HCOLOR 1,3 PARENT oBrowse   ;
		PICTURE '99999';
	   EDITPROTECT{|| TRUE }

   DCBROWSECOL FIELD UCAPPRSL->LINES                     ;
		HEADER "LINES" HCOLOR 1,3 PARENT oBrowse   ;
		PICTURE '9';
	   EDITPROTECT{|| TRUE }


	DCBROWSECOL FIELD UCAPPRSL->VIN                     ;
		HEADER "**VIN" HCOLOR 1,3 PARENT oBrowse   ;
		SORT{|| cSearch:='Appraisal Vin     ',UCAPPRSL->(ORDSETFOCUS('UCAPPVIN')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)} ;
	   EDITPROTECT{|| TRUE }

	DCBROWSECOL FIELD UCAPPRSL->SM                     ;
		HEADER "**SM" HCOLOR 1,3 PARENT oBrowse   ;
		SORT{|| UCAPPRSL->(ORDSETFOCUS('UCAPPSM')),dbgotop(),oBrowse:REFRESHALL()} ;
	   EDITPROTECT{|| TRUE }



	DCREAD GUI ;
		FIT ;
		TITLE 'Used Vehicle Appraisal Browser' ;
		ADDBUTTONS   ;
		OPTIONS GETOPTIONS

	CLOSE ALL
ReTURN nil
*** END OF EXAMPLE ***


STATIC FUNCTION _FindUpRec(cUpid,oDeskRecord)
	select NEWUPS
	ordsetfocus('NEWUPIN')
	SEEK cUpid
	IF found()                      /// make record object
		oDeskRecord:=NEWUPS->(DC_DBRECORD(): NEW())
		NEWUPS->(DC_DBSCATTER(oDeskRecord))
		RETURN TRUE
	ELSE
		DC_WINALERT('The Up Record For this Appraisal Is No Longer Available')
	ENDIF
	RETURN FALSE


FUNCTION SBROWBYDESC(cDemodStkNo)
	LOCAL GetList := {}, aPres, oBrowse, oToolBar , getoptions, cFindme, clAvailcode,cIndexOrder:='Model Description Order ',cSearch:=space(20)
	LOCAL XSTATUS:=.T. ,oDialog
	LOCAL cProgram , nRec

	nREC:=0
	cProgram:='BROWBYDE'
	SELECT NCI
	NCI->(ordsetfocus('invidx03'))
	NCI->(DBGOTOP())
	DCGETOPTIONS ;
		SAYWIDTH 0 ;
		AUTORESIZE
	SET FILTER TO NCI->AVAILCODE='A'

	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

@ 1,1 DCSAY 'Current Index Order' get cIndexOrder EDITPROTECT{||.T.}
@ 1,70 DCSAY 'Search For Key' GET cSearch SAYRIGHT PICT '@!' ;
      KEYBLOCK { |a,b,o| DBSELECTAREA('NCI'),DC_BrowseAutoSeek(a,o,oBrowse)} ;
		GETID 'Srchkey'

@ 2,1 DCSAY 'You may RIGHT CLICK on Column Headers that Begin with ** to change Search Order to that Column '


/* ----- Create browse ----- */

	@ 3.5,1 DCBROWSE oBrowse ALIAS 'NCI'                 ;
		SIZE 130,20                                     ;
		PRESENTATION aPres;
		FREEZELEFT {1,2,3,4}     ;
		SORTSCOLOR 0,2;
		SORTUCOLOR 1,6  ;
		DATALINK{|| cDemodStkNo:=NCI->STOCKNO,DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST) }

	DCBROWSECOL FIELD NCI->STOCKNO                     ;
		HEADER "**STOCK#  " HCOLOR 9,3 PARENT oBrowse                  ;
		SORT{|| oDialog:=DC_WAITON('Now Sorting by Stock Number'),cIndexOrder:='Stock Number Order          ',DBSELECTAREA('NCI'),;
		NCI->(ordsetfocus('INVIDX01')),;
		DC_IMPL(oDialog),DC_DbGoTop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)}


	DCBROWSECOL FIELD NCI->LAST6 ;
		HEADER "**Last 6 " HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKRED PARENT oBrowse ;
		WIDTH 6 ;
		SORT{|| oDialog:=DC_WAITON('Now Sorting by Last 6'),cIndexOrder:='Last 6 Of Vin           ',DBSELECTAREA('NCI'),;
		NCI->(ordsetfocus('INVLST6')),;
		DC_IMPL(oDialog),DC_DbGoTop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)}


	DCBROWSECOL FIELD NCI->YEAR ;
		HEADER "YEAR" PARENT oBrowse

	DCBROWSECOL FIELD NCI->MAKE ;
		HEADER "**MAKE" WIDTH 10 HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse ;
		SORT{|| oDialog:=DC_WAITON('Now Sorting by Make'),cIndexOrder:='Make               ',DBSELECTAREA('NCI'),;
		NCI->(ordsetfocus('INVMAKE')),;
		dc_impl(oDialog),;
		DC_DbGoTop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)}

	DCBROWSECOL FIELD NCI->MODEL ;
		HEADER "**MODEL" WIDTH 10 HCOLOR GRA_CLR_WHITE,GRA_CLR_DARKGREEN PARENT oBrowse  ;
		SORT{|| oDialog:=DC_WAITON('Now Sorting by Description'),cIndexOrder:='Model Description',DBSELECTAREA('NCI'),;
		NCI->(ordsetfocus('INVIDX03')),;
		dc_impl(oDialog),;
		DC_DbGoTop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST)}

	DCBROWSECOL FIELD NCI->AVAILCODE ;
		HEADER "AVAILCODE" PARENT oBrowse
	DCBROWSECOL FIELD NCI->retail ;
		HEADER "MSRP" PARENT oBrowse
	DCBROWSECOL FIELD NCI->cost ;
		HEADER "TISSUE" PARENT oBrowse     ;
		HIDE {|| IF(!MANAGER(),.T.,.F.)}
	DCBROWSECOL FIELD NCI->trim ;
		HEADER "SERIES" PARENT oBrowse
	DCBROWSECOL FIELD NCI->BODY ;
		HEADER "BODYTYPE" PARENT oBrowse
	DCBROWSECOL FIELD NCI->ENGINE ;
		HEADER "ENGINE" PARENT oBrowse

	DCBROWSECOL FIELD NCI->TRANSmiss ;
		HEADER "TRANS" PARENT oBrowse

	DCBROWSECOL FIELD NCI->COLOR ;
		HEADER "COLOR" PARENT oBrowse
	DCBROWSECOL FIELD NCI->INTCODE ;
		HEADER "INTCODE" PARENT oBrowse
	DCBROWSECOL FIELD NCI->VIN ;
		HEADER "VIN" PARENT oBrowse;
		PROTECT {|| .T.}
	DCBROWSECOL FIELD NCI->REC_FROM ;
		HEADER "RECEIVED FROM " PARENT oBrowse

	DCBROWSECOL FIELD NCI->PKG1 ;
		HEADER "PKG1" PARENT oBrowse
	DCBROWSECOL FIELD NCI->PKG2 ;
		HEADER "PKG2" PARENT oBrowse
	DCBROWSECOL FIELD NCI->PKG3 ;
		HEADER "PKG3" PARENT oBrowse
	DCBROWSECOL FIELD NCI->OPT1 ;
		HEADER "OPTION" PARENT oBrowse
	DCBROWSECOL FIELD NCI->OPT2 ;
		HEADER "OPTION" PARENT oBrowse
	DCBROWSECOL FIELD NCI->OPT3 ;
		HEADER "OPTION" PARENT oBrowse
	DCBROWSECOL FIELD NCI->OPT4 ;
		HEADER "OPTION" PARENT oBrowse
	DCBROWSECOL FIELD NCI->OPT5 ;
		HEADER "OPTION" PARENT oBrowse
	DCBROWSECOL FIELD NCI->OPT6 ;
		HEADER "OPTION" PARENT oBrowse
	DCBROWSECOL FIELD NCI->customer ;
		HEADER "SOLDTO" PARENT oBrowse
	DCBROWSECOL FIELD NCI->stkDATE ;
		HEADER "DATE IN" PARENT oBrowse
	DCREAD GUI ;
		FIT ;
		TO xStatus ;
		OPTIONS GETOPTIONS ;
		TITLE 'NEW VEHICLE BY DESCRIPTION';
		ADDBUTTONS

  /// if no choice is made return blank
  IF !xStatus
	  cDemodStkNo:=space(8)
	  RETURN NIL
  ENDIF


	cDemodStkNo:=NCI->STOCKNO


ReTURN cDemodStkNo
*** END OF EXAMPLE ***

FUNCTION BROWBYtype()
	LOCAL GETLIST:={}, getoptions, xStatus , aPres, oRecord
	LOCAL nType:=1 , dDatex , oBrowse , Aotoolbar
	local cProgram:='BROWBYtyp' , lNoShowDead:=.t.
	IF.NOT.SECURITY(@cProgram)
	 DC_winaLERT('SECURITY VIOLATION')
    RETURN NIL
   ENDIF

IF !OPENUPSFILES()
	CLOSE ALL
   RETURN NIL
ENDIF
NEWUPS->(ORDSETFOCUS('NEWUPTYP'))

DCGETOPTIONS ;
	SAYWIDTH 0 ;
	AUTORESIZE
XSTATUS:=.T.
dDATEX:=(DATE()-7)




@ 10,0 DCSAY 'Enter the Type you wish to Browse for 1,2,3,4' get nType range 1,4 picture '9'  SAYCOLOR 3,0 SAYFONT '10.COURIER' GETFONT '12.COURIER'
//@ 14,0 DCCHECKBOX lNoShowDead PROMPT 'Click here to NOT show deals already DEAD'

DCREAD GUI TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Up Source Browser'
IF !XSTATUS
 RETURN NIL
ENDIF
SELECT NEWUPS
ordsetfocus('newuptyp')
/// now get to date requested
SEEK nType

do while NEWUPS->datein < ddatex
	skip alias newups
enddo

IF lNoShowDead
	SET FILTER TO NEWUPS->UPSTATUS <> 'D'
ENDIF



aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

@ 1,1 DCTOOLBAR AoToolBar                               ;
	SIZE 120, 1.5



DCADDBUTTON CAPTION 'Desking Tool;Quote '                              ;
	;//ACTION {||oRecord:=NEWUPS->(DC_DBRECORD():NEW()),;
	;//NEWUPS->(DC_DBSCATTER(oRecord)),;
	ACTION {|| UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)};
	PARENT AoToolBar


DCADDBUTTON CAPTION '2.New Vehicle;Browser  '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("2")};
	ACTION {||SBROWBYDESC(),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar

DCADDBUTTON CAPTION '3.Used Vehicle;Browser '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("3")};
	ACTION {||SUBROWSTK(),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar

DCADDBUTTON CAPTION '4.Maintain Up;Record '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("4")};
	ACTION {|| UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar

DCADDBUTTON CAPTION '5.COMMENTS '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("5")};
	ACTION {|| UPNOTES(),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar

DCADDBUTTON CAPTION '6.Print UPSheet '                              ;
	SIZE 20                                              ;
	ACCELKEY {ASC("6")};
	ACTION {||PRINTUPSHEET(NEWUPS->(RECNO()))};
	PARENT AOToolBar

DCADDBUTTON CAPTION '7.Send eMail '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("7")};
	ACTION {||upemail()};
	PARENT aoToolBar


@ 3,0 DCSAY 'COLOR CODES'
@ 3,15 DCSAY 'Sold' SAYCOLOR 1,5
@ 3,30 DCSAY 'Active' SAYCOLOR 1,0
@ 3,45 DCSAY 'Tickle' SAYCOLOR 1,15
@ 3,60 DCSAY 'DEAD' SAYCOLOR 0,1
@ 3,75 DCSAY 'Rescue' SAYCOLOR 2,3
@ 3,90 DCSAY 'PhoneUp' SAYCOLOR 12,7
@ 3,105 DCSAY 'InterNetUp' SAYCOLOR 1,4


/* ----- Create browse ----- */

@ 5,1 DCBROWSE oBrowse ALIAS 'NEWUPS'                 ;
	SIZE 110,20                                     ;
	PRESENTATION aPres;
	FREEZELEFT {1,2,3};
	color{|| getupcolor(NEWUPS->UPSTATUS)};
	EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITACROSS
	//DATALINK{||MAINUP(.T.)}


DCBROWSECOL FIELD NEWUPS->Salesman ;
	HEADER "SalesPer" PARENT oBrowse  ;
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->Upstatus ;
	HEADER "Status" PARENT oBrowse  ;
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1

DCBROWSECOL FIELD NEWUPS->ORIGSOURCE ;
	HEADER "Source" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1


DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "UPID/PHN#" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 7

DCBROWSECOL FIELD NEWUPS->PREFMETHOD ;
	HEADER "PREFMETHOD" PARENT oBrowse  ;
	WIDTH 4



DCBROWSECOL FIELD NEWUPS->LASTNAME                     ;
	HEADER "LASTNAME" HCOLOR 9,3 PARENT oBrowse     ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->FIRSTNAME ;
	HEADER "FIRSTNAME" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->DATEIN ;
	HEADER "DATEIN" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->EMAIL ;
	HEADER "EMAIL" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->STREET ;
	HEADER "STREET" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->CITYSTATE ;
	HEADER "CITYSTATE" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->ZIPCODE ;
	HEADER "ZIPCODE" PARENT oBrowse  ;
	WIDTH 5
DCBROWSECOL FIELD NEWUPS->AREA ;
	HEADER "AREACODE" PARENT oBrowse  ;
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->WORKPHN ;
	HEADER "WORKPHN" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->cellphone ;
	HEADER "CELLPHN" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->beback1 ;
	HEADER "Beback1" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->beback2 ;
	HEADER "Beback2" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->NEWUSED ;
	HEADER "NEWUSED" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->INTEREST ;
	HEADER "INTEREST" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->SALECODE ;
	HEADER "SALECODE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->TOMGR ;
	HEADER "TOMGR" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->COMMENTS ;
	HEADER "COMMENT" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->COMMENT2 ;
	HEADER "COMMENT2" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->ADSOURCE ;
	HEADER "ADSOURCE" PARENT oBrowse  ;
	WIDTH 4
DCBROWSECOL FIELD NEWUPS->DOWNDEAL ;
	HEADER "DOWN" PARENT oBrowse  ;
	WIDTH 1

DCREAD GUI ;
	FIT ;
	EVAL {||SETAPPFOCUS(obROWSE:GETCOLUMN(1))};
	BUTTONS DCGUI_BUTTON_EXIT
ReTURN nil


FUNCTION BROWBYSOURCE()
	LOCAL GETLIST:={}, getoptions, xStatus , aPres
	LOCAL cSrc:='F' , dDatex , oBrowse , Aotoolbar,oRecord
	local cProgram:='BROWBYtyp' , lNoShowDead:=.t.
	IF.NOT.SECURITY(@cProgram)
	 DC_winaLERT('SECURITY VIOLATION')
    RETURN NIL
   ENDIF

IF !OPENUPSFILES()
	CLOSE ALL
   RETURN NIL
ENDIF
NEWUPS->(ORDSETFOCUS('NEWUPSRC'))

DCGETOPTIONS ;
	SAYWIDTH 0 ;
	AUTORESIZE
XSTATUS:=.T.
dDATEX:=(DATE()-7)


@ 10,0 DCSAY 'Enter the Type you wish to Browse for F I P ' get cSrc  picture '@!' VALID{|| IIF( cSrc $('FIP'),.t. ,.f. )} SAYCOLOR 3,0 SAYFONT '10.COURIER' GETFONT '12.COURIER'
//@ 14,0 DCCHECKBOX lNoShowDead PROMPT 'Click here to NOT show deals already DEAD'
DCREAD GUI TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Up Source Browser'
IF !XSTATUS
 RETURN NIL
ENDIF
SELECT NEWUPS
ordsetfocus('newupsrc')
/// now get to date requested

SEEK cSrc

do while NEWUPS->datein < ddatex
	skip alias newups
enddo

IF lNoShowDead
	SET FILTER TO NEWUPS->upstatus <> 'D'
ENDIF



aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

@ 1,1 DCTOOLBAR AoToolBar                               ;
	SIZE 120, 1.5

DCADDBUTTON CAPTION 'Desking Tool;Quote '                              ;
	;//ACTION {||oRecord:=NEWUPS->(DC_DBRECORD():NEW()),;
	;//NEWUPS->(DC_DBSCATTER(oRecord)),;
	ACTION {|| UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)};
	PARENT AoToolBar



DCADDBUTTON CAPTION '2.New Vehicle;Browser  '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("2")};
	ACTION {||SBROWBYDESC(),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar

DCADDBUTTON CAPTION '3.Used Vehicle;Browser '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("3")};
	ACTION {||SUBROWSTK(),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar

DCADDBUTTON CAPTION '4.Maintain Up;Record '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("4")};
	ACTION {|| UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar

DCADDBUTTON CAPTION '5.COMMENTS '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("5")};
	ACTION {|| UPNOTES(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(obROWSE:GETCOLUMN(1))  };
	PARENT AoToolBar

DCADDBUTTON CAPTION '6.Print UPSheet '                              ;
	SIZE 20                                              ;
	ACCELKEY {ASC("6")};
	ACTION {||PRINTUPSHEET(NEWUPS->(RECNO()))};
	PARENT AOToolBar

DCADDBUTTON CAPTION '7.Send eMail '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("7")};
	ACTION {||upemail()};
	PARENT aoToolBar


@ 3,0 DCSAY 'COLOR CODES'
@ 3,15 DCSAY 'Sold' SAYCOLOR 1,5
@ 3,30 DCSAY 'Active' SAYCOLOR 1,0
@ 3,45 DCSAY 'Tickle' SAYCOLOR 1,15
@ 3,60 DCSAY 'DEAD' SAYCOLOR 0,1
@ 3,75 DCSAY 'Rescue' SAYCOLOR 2,3
@ 3,90 DCSAY 'PhoneUp' SAYCOLOR 12,7
@ 3,105 DCSAY 'InterNetUp' SAYCOLOR 1,4



/* ----- Create browse ----- */

@ 5,1 DCBROWSE oBrowse ALIAS 'NEWUPS'                 ;
	SIZE 110,20                                     ;
	PRESENTATION aPres;
	FREEZELEFT {1,2,3};
	color{|| getupcolor(NEWUPS->UPSTATUS)};
	EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITACROSS ;
	//DATALINK{||MAINUP(.T.)}


DCBROWSECOL FIELD NEWUPS->Salesman ;
	HEADER "SalesPer" PARENT oBrowse  ;
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->Upstatus ;
	HEADER "Status" PARENT oBrowse  ;
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1

DCBROWSECOL FIELD NEWUPS->ORIGSOURCE ;
	HEADER "Source" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1


DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "UPID/PHN#" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 7

DCBROWSECOL FIELD NEWUPS->PREFMETHOD ;
	HEADER "PREFMETHOD" PARENT oBrowse  ;
	WIDTH 4


DCBROWSECOL FIELD NEWUPS->LASTNAME                     ;
	HEADER "LASTNAME" HCOLOR 9,3 PARENT oBrowse     ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->FIRSTNAME ;
	HEADER "FIRSTNAME" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->DATEIN ;
	HEADER "DATEIN" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->EMAIL ;
	HEADER "EMAIL" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->STREET ;
	HEADER "STREET" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->CITYSTATE ;
	HEADER "CITYSTATE" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->ZIPCODE ;
	HEADER "ZIPCODE" PARENT oBrowse  ;
	WIDTH 5
DCBROWSECOL FIELD NEWUPS->AREA ;
	HEADER "AREACODE" PARENT oBrowse  ;
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->WORKPHN ;
	HEADER "WORKPHN" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->cellphone ;
	HEADER "CELLPHN" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->beback1 ;
	HEADER "Beback1" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->beback2 ;
	HEADER "Beback2" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->NEWUSED ;
	HEADER "NEWUSED" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->INTEREST ;
	HEADER "INTEREST" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->SALECODE ;
	HEADER "SALECODE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->TOMGR ;
	HEADER "TOMGR" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->COMMENTS ;
	HEADER "COMMENT" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->COMMENT2 ;
	HEADER "COMMENT2" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->ADSOURCE ;
	HEADER "ADSOURCE" PARENT oBrowse  ;
	WIDTH 4
DCBROWSECOL FIELD NEWUPS->DOWNDEAL ;
	HEADER "DOWN" PARENT oBrowse  ;
	WIDTH 1

DCREAD GUI ;
	FIT ;
	EVAL {||SETAPPFOCUS(obROWSE:GETCOLUMN(1))};
	BUTTONS DCGUI_BUTTON_EXIT
ReTURN nil



FUNCTION OPENUPSFILES()

IF !UseDb({'INETUPS','NCI','ZIPPER','USAZIP','UCL01','NEWUPS','UPAPPTS','CLASSORT','ADSOURCE','SM','SERVCUST','UCAPPRSL','EMAILLOG','REBATEUP','USEDOPTS','ACELEASE','ACEBANK'})
	 DC_WINALERT('Cannot Open Files     .')
    RETURN FALSE
ENDIF
RETURN TRUE

static function _IndexAvvSync()
	LOCAL aStruct:={}
	IF !fexists('AVVSYNC.DBF')
	  aStruct:={}
	  AADD(aStruct,{"LASTNAME","C",30,0})
	  AADD(aStruct,{"FIRSTNAME","C",30,0})
	  AADD(aStruct,{"STREET","C",30,0})
	  AADD(aStruct,{"CITY","C",30,0})
	  AADD(aStruct,{"FULLNAME","C",50,0})
	  AADD(aStruct,{"ID","C",10,0})
	  AADD(aStruct,{"MIDDLE","C",10,0})
	  AADD(aStruct,{"DAYPHONE","C",15,0})
	  AADD(aStruct,{"DEALER","C",10,0})
	  AADD(aStruct,{"EMAIL","C",50,0})
	  AADD(aStruct,{"EVEPHON","C",10,0})
	  AADD(aStruct,{"GROUPDIR","C",10,0})
	  AADD(aStruct,{"LASTACT","C",10,0})
	  AADD(aStruct,{"MAKE","C",15,0})
	  AADD(aStruct,{"MODEL","C",10,0})
	  AADD(aStruct,{"ZIPCODE","C",5,0})
	  AADD(aStruct,{"PREFCONT","C",10,0})
	  AADD(aStruct,{"SALESCYS","C",10,0})
	  AADD(aStruct,{"SALESID","C",10,0})
	  AADD(aStruct,{"CONFIRST","C",10,0})
	  AADD(aStruct,{"CONLAST","C",10,0})
	  AADD(aStruct,{"SOLDID","C",10,0})
	  AADD(aStruct,{"SOLDTYPE","C",10,0})
	  AADD(aStruct,{"SOLDTIME","C",10,0})
	  AADD(aStruct,{"SOURCE","C",10,0})
	  AADD(aStruct,{"STATE","C",10,0})
	  AADD(aStruct,{"STOCKNO","C",10,0})
	  AADD(aStruct,{"TIMESTAMP","C",10,0})
	  AADD(aStruct,{"VEHOTH","C",10,0})
	  AADD(aStruct,{"YEAR","C",10,0})
	  AADD(aStruct,{"OPTOUT","C",10,0})
	  AADD(aStruct,{"SM","C",2,0})
	  DBCREATE('AVVSYNC',aStruct,'FOXCDX')
  ENDIF

	/*
	IF USE_UDF('AVVSYNC',.T.)
		ORDCREATE('AVVSYNC','AVVSYNC','EMAIL')
		ORDCREATE('AVVLAST','AVVLAST','UPPER(LASTNAME)')
		ORDCREATE('AVVSM','AVVSM','CONLAST')
	ENDIF    */
	///AVVSYNC->(DBCLOSEAREA())
	RETURN nil

static function _IndexInetups()
	local oDialog ,aStruct:={}
	IF !fexists('INETUPS.DBF')
	  aStruct:={}
	  AADD(aStruct,{"LASTNAME","C",30,0})
	  AADD(aStruct,{"FIRSTNAME","C",30,0})
	  AADD(aStruct,{"STREET","C",30,0})
	  AADD(aStruct,{"CITY","C",30,0})
	  AADD(aStruct,{"FULLNAME","C",50,0})
	  AADD(aStruct,{"ID","C",10,0})
	  AADD(aStruct,{"MIDDLE","C",10,0})
	  AADD(aStruct,{"DAYPHONE","C",15,0})
	  AADD(aStruct,{"DEALER","C",10,0})
	  AADD(aStruct,{"EMAIL","C",50,0})
	  AADD(aStruct,{"EVEPHON","C",15,0})
	  AADD(aStruct,{"GROUPDIR","C",10,0})
	  AADD(aStruct,{"LASTACT","C",10,0})
	  AADD(aStruct,{"MAKE","C",15,0})
	  AADD(aStruct,{"MODEL","C",10,0})
	  AADD(aStruct,{"ZIPCODE","C",5,0})
	  AADD(aStruct,{"PREFCONT","C",10,0})
	  AADD(aStruct,{"SALESCYS","C",10,0})
	  AADD(aStruct,{"SALESID","C",10,0})
	  AADD(aStruct,{"CONFIRST","C",10,0})
	  AADD(aStruct,{"CONLAST","C",10,0})
	  AADD(aStruct,{"SOLDID","C",10,0})
	  AADD(aStruct,{"SOLDTYPE","C",10,0})
	  AADD(aStruct,{"SOLDTIME","C",10,0})
	  AADD(aStruct,{"SOURCE","C",10,0})
	  AADD(aStruct,{"STATE","C",10,0})
	  AADD(aStruct,{"STOCKNO","C",10,0})
	  AADD(aStruct,{"TIMESTAMP","C",10,0})
	  AADD(aStruct,{"VEHOTH","C",10,0})
	  AADD(aStruct,{"YEAR","C",10,0})
	  AADD(aStruct,{"OPTOUT","C",10,0})
	  AADD(aStruct,{"DATERECVD","D",8,0})
	  AADD(aStruct,{"SM","C",2,0})
	  AADD(aStruct,{"COMMENT","M",10,0})



	  DBCREATE('INETUPS',aStruct,'FOXCDX')
  ENDIF
  IF !UseDb({'INETUPS'},1,.F.,.T.)
  	 DC_WINALERT('Cannot Open Files     .')
     	 RETURN NIL
  ENDIF

	INETUPS->(DBCLOSEAREA())
	RETURN nil




static function _IndexUpAppts()
	LOCAL aStruct:={}
	IF !FEXISTS('UPAPPTS.DBF')
	 aStruct:={}
	 AADD(aStruct,{"UPID","C",7,0})
	 AADD(aStruct,{"APPTDATE","D",8,0})
	 AADD(aStruct,{"REASON","C",120,0})
	 AADD(aStruct,{"SHOWED","L",1,0})
	 AADD(aStruct,{"EMAIL","C",50,0})
	 AADD(aStruct,{"LASTNAME","C",30,0})
	 AADD(aStruct,{"FIRSTNAME","C",50,0})
	 AADD(aStruct,{"PHONE","C",10,0})
	 AADD(aStruct,{"APPTTIME","C",10,0})
	 AADD(aStruct,{"SM","C",2,0})
	 AADD(aStruct,{"FIRSTAPPT","L",1,0})
	 AADD(aStruct,{"FOLLOWAPPT","L",1,0})
	 AADD(aStruct,{"BLANKCHAR","C",10,0})
	 AADD(aStruct,{"ORIGDATE","D",8,0})
	 AADD(aStruct,{"UPSOURCE","C",1,0})
	 AADD(aStruct,{"UPTYPE","N",1,0})
	 AADD(aStruct,{"NEWUSED","C",1,0})
	 AADD(aStruct,{"INTEREST","C",10,0})
	 DBCREATE('UPAPPTS',aStruct,'FOXDCX')
	ENDIF
	IF !UseDb({'UPAPPTS'})
	 	 DC_WINALERT('Cannot Open Files     .')
	    RETURN NIL
	ENDIF

	UPAPPTS->(DBCLOSEAREA())
	RETURN NIL


FUNCTION BROWBYNM()
	LOCAL GETLIST:={} , aPres  , oBrowse , Aotoolbar , getoptions , xstatus ,oRecord
	LOCAL cLastName:=SPACE(40) , dDatex:=DATE()-60
	local cProgram:='BROWBYNM'  , lNoShowDead:=.t.
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
	IF.NOT.SECURITY(@cProgram)
	 DC_winaLERT('SECURITY VIOLATION')
    RETURN NIL
   ENDIF

/// OPEN ALL FILES
IF !OPENUPSFILES()
	CLOSE ALL
   RETURN NIL
ENDIF



NEWUPS->(ORDSETFOCUS('NEWUPNM'))
NEWUPS->(DBGOTOP())

@ 10,0 DCSAY 'Enter The Lastname You Wish To Browse' GET cLastName SAYCOLOR 3,0 SAYFONT '10.COURIER' GETFONT '12.COURIER'


DCREAD GUI TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'ENTER LASTNAME TO BEGIN BROWSE'
IF !XSTATUS
 RETURN NIL
ENDIF
SELECT NEWUPS
SET SOFTSEEK ON
SEEK UPPER(cLastName)
SET SOFTSEEK OFF

@ 3,0 DCSAY 'COLOR CODES'
@ 3,15 DCSAY 'Sold' SAYCOLOR 1,5
@ 3,30 DCSAY 'Active' SAYCOLOR 1,0
@ 3,45 DCSAY 'Tickle' SAYCOLOR 1,15
@ 3,60 DCSAY 'DEAD' SAYCOLOR 0,1
@ 3,75 DCSAY 'Rescue' SAYCOLOR 2,3
@ 3,90 DCSAY 'PhoneUp' SAYCOLOR 12,7
@ 3,105 DCSAY 'InterNetUp' SAYCOLOR 1,4



aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */
@ 5,1 DCTOOLBAR AoToolBar                               ;
	SIZE 132, 1.5

IF MANAGER()
 DCADDBUTTON CAPTION 'Desking Tool;Quote '                              ;
	;//ACTION {||oRecord:=NEWUPS->(DC_DBRECORD():NEW()),;
	;//NEWUPS->(DC_DBSCATTER(oRecord)),;
	ACTION {|| UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)};
	PARENT AoToolBar
ENDIF

DCADDBUTTON CAPTION '2.New Vehicle;Browser  '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("2")};
	ACTION {||SBROWBYDESC(),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar

DCADDBUTTON CAPTION '3.Used Vehicle;Browser '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("3")};
	ACTION {||SBROWBYDESC(),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar

DCADDBUTTON CAPTION '4.Maintain Up;Record '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("4")};
	ACTION {|| UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar


DCADDBUTTON CAPTION '5.COMMENTS '                              ;
	SIZE 12                                              ;
	ACCELKEY {ASC("5")};
	ACTION {|| UPNOTES(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(obROWSE:GETCOLUMN(1))  };
	PARENT AoToolBar
DCADDBUTTON CAPTION '6.Print UPSheet '                              ;
	SIZE 20                                              ;
	ACCELKEY {ASC("6")};
	ACTION {||PRINTUPSHEET(NEWUPS->(RECNO()))};
	PARENT aoToolBar

DCADDBUTTON CAPTION '7.Send eMail '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("7")};
	ACTION {||upemail()};
	PARENT aoToolBar

DCADDBUTTON CAPTION '8.Review eMails '                              ;
	   ACCELKEY {ASC("8")};
		SIZE 15                                              ;
		ACTION {|| BROWEMAILLOG('',NEWUPS->EMAIL)} ;
		PARENT AoToolBar


/* ----- Create browse ----- */

@ 8,1 DCBROWSE oBrowse ALIAS 'NEWUPS'                 ;
	SIZE 110,20                                     ;
	PRESENTATION aPres;
	FREEZELEFT {1,2,3};
	color{|| getupcolor(NEWUPS->UPSTATUS)};
	EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITACROSS ;
	//DATALINK{||MAINUP(.T.)}

DCBROWSECOL FIELD NEWUPS->Salesman ;
	HEADER "SalesPer" PARENT oBrowse  ;
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->Upstatus ;
	HEADER "Status" PARENT oBrowse  ;
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "UPID/PHN#" PARENT oBrowse  ;
	WIDTH 7

DCBROWSECOL FIELD NEWUPS->PREFMETHOD ;
	HEADER "PREFMETHOD" PARENT oBrowse  ;
	WIDTH 4


DCBROWSECOL FIELD NEWUPS->SALESMAN ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->LASTNAME                     ;
	HEADER "LASTNAME" HCOLOR 9,3 PARENT oBrowse     ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->FIRSTNAME ;
	HEADER "FIRSTNAME" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->DATEIN ;
	HEADER "DATEIN" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->EMAIL ;
	HEADER "EMAIL" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->STREET ;
	HEADER "STREET" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->CITYSTATE ;
	HEADER "CITYSTATE" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->ZIPCODE ;
	HEADER "ZIPCODE" PARENT oBrowse  ;
	WIDTH 5
DCBROWSECOL FIELD NEWUPS->AREA ;
	HEADER "AREACODE" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "PHONE" PARENT oBrowse  ;
	WIDTH 7
DCBROWSECOL FIELD NEWUPS->WORKPHN ;
	HEADER "WORKPHN" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->cellphone ;
	HEADER "CELLPHN" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->beback1 ;
	HEADER "Beback1" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->beback2 ;
	HEADER "Beback2" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->NEWUSED ;
	HEADER "NEWUSED" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->ORIGSOURCE ;
	HEADER "SOURCE" PARENT oBrowse  ;
	WIDTH 1


DCBROWSECOL FIELD NEWUPS->INTEREST ;
	HEADER "INTEREST" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->SALECODE ;
	HEADER "SALECODE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->TOMGR ;
	HEADER "TOMGR" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->COMMENTS ;
	HEADER "COMMENT" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->COMMENT2 ;
	HEADER "COMMENT2" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->ADSOURCE ;
	HEADER "ADSOURCE" PARENT oBrowse  ;
	WIDTH 4
DCBROWSECOL FIELD NEWUPS->DOWNDEAL ;
	HEADER "DOWN" PARENT oBrowse  ;
	WIDTH 1

DCREAD GUI ;
	FIT ;
	OPTIONS GETOPTIONS ;
	EVAL {||SETAPPFOCUS(obROWSE:GETCOLUMN(1))};
	BUTTONS DCGUI_BUTTON_EXIT ;
	TITLE 'Automan Up Browser by Name'
ReTURN nil


FUNCTION BROWUPS(nREC)
	LOCAL GETLIST:={} , xStatus, AoToolbar , oBrowse ,aPres , getoptions
	LOCAL cLastName:=SPACE(25)
	LOCAL cProgram:='BROWUPS'
	IF.NOT.SECURITY(@cProgram)
   	DC_WINALERT('SECURITY VIOLATION')
      RETURN NIL
   ENDIF

DCGETOPTIONS ;
	SAYWIDTH 0 ;
	AUTORESIZE
   XSTATUS:=.T.
aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

@ 1,1 DCTOOLBAR AoToolBar                               ;
	SIZE 120, 1.5

@ 3,0 DCSAY 'COLOR CODES'
@ 3,15 DCSAY 'Sold' SAYCOLOR 1,BD_KEYLIME
@ 3,30 DCSAY 'Active' SAYCOLOR 1, BD_WASHGREY
@ 3,45 DCSAY 'Tickle' SAYCOLOR 1,BD_CORNFLOWERBLUE
@ 3,60 DCSAY 'DEAD' SAYCOLOR 1,BD_SALMON
@ 3,75 DCSAY 'Rescue' SAYCOLOR 2,BD_SUNNYYELLOW
@ 3,90 DCSAY 'PhoneUp' SAYCOLOR 2,BD_FADEDPURPLE
@ 3,105 DCSAY 'InterNetUp' SAYCOLOR 1,BD_BLUEMIST



/* ----- Create browse ----- */

@ 5,1 DCBROWSE oBrowse ALIAS 'NEWUPS'                 ;
	SIZE 110,20                                     ;
	PRESENTATION aPres;
	FREEZELEFT {1,2,3};
	color{|| getupcolor(NEWUPS->UPSTATUS)};
	EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITACROSS ;
	DATALINK{||nREC:=NEWUPS->(RECNO()) }

DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1

DCBROWSECOL FIELD NEWUPS->ORIGSOURCE COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)};
	HEADER "SOURCE" PARENT oBrowse  ;
	WIDTH 1

DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "UPID/PHN#" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 7

DCBROWSECOL FIELD NEWUPS->PREFMETHOD ;
	HEADER "PREFMETHOD" PARENT oBrowse  ;
	WIDTH 4


DCBROWSECOL FIELD NEWUPS->SALESMAN ;
	HEADER "SALESMAN" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->LASTNAME                     ;
	HEADER "LASTNAME" HCOLOR 9,3 PARENT oBrowse     ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->FIRSTNAME ;
	HEADER "FIRSTNAME" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->DATEIN ;
	HEADER "DATEIN" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->beback1 ;
	HEADER "Beback1" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->beback2 ;
	HEADER "Beback2" PARENT oBrowse  ;
	WIDTH 8

DCBROWSECOL FIELD NEWUPS->EMAIL ;
	HEADER "EMAIL" PARENT oBrowse  ;
	WIDTH 35

DCBROWSECOL FIELD NEWUPS->STREET ;
	HEADER "STREET" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->CITYSTATE ;
	HEADER "CITYSTATE" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->ZIPCODE ;
	HEADER "ZIPCODE" PARENT oBrowse  ;
	WIDTH 5
DCBROWSECOL FIELD NEWUPS->AREA ;
	HEADER "AREACODE" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "PHONE" PARENT oBrowse  ;
	WIDTH 7
DCBROWSECOL FIELD NEWUPS->WORKPHN ;
	HEADER "WORKPHN" PARENT oBrowse  ;
	WIDTH 12
DCBROWSECOL FIELD NEWUPS->cellphone ;
	HEADER "CELLPHN" PARENT oBrowse  ;
	WIDTH 10

DCBROWSECOL FIELD NEWUPS->EMAIL ;
	HEADER "EMAIL" PARENT oBrowse  ;
	WIDTH 35

DCBROWSECOL FIELD NEWUPS->NEWUSED ;
	HEADER "NEWUSED" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->INTEREST ;
	HEADER "INTEREST" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->SALECODE ;
	HEADER "SALECODE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->TOMGR ;
	HEADER "TOMGR" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->COMMENTS ;
	HEADER "COMMENT" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->COMMENT2 ;
	HEADER "COMMENT2" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->ADSOURCE ;
	HEADER "ADSOURCE" PARENT oBrowse  ;
	WIDTH 4
DCBROWSECOL FIELD NEWUPS->DOWNDEAL ;
	HEADER "DOWN" PARENT oBrowse  ;
	WIDTH 1
DCREAD GUI ;
	FIT ;
	TO xStatus ;
	OPTIONS GETOPTIONS ;
	EVAL {||SETAPPFOCUS(obROWSE:GETCOLUMN(1))};
	ADDBUTTONS
   IF !xStatus
	 nRec:=0
	 RETURN nRec
   ENDIF
   nRec:=NEWUPS->(RECNO())

ReTURN nREC

FUNCTION BROWINETUPS(nREC)
	LOCAL GETLIST:={} , xStatus, oToolbar , oBrowse ,aPres , getoptions, cIndexorder:='Last Name Order    ',  cSearchkey:=space(20)
	LOCAL cLastName:=SPACE(25) ,oRecord
	LOCAL cProgram:='BROWUPS'
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '8.Courier New' GETFONT '8.Courier New' AUTORESIZE
	set deleted on
	IF.NOT.SECURITY(@cProgram)
   	DC_WINALERT('SECURITY VIOLATION')
      RETURN NIL
   ENDIF

IF !OPENUPSFILES()
	RETURN NIL
ENDIF
DCGETOPTIONS ;
	SAYWIDTH 0 ;
	AUTORESIZE


select INETUPS
INETUPS->(ORDSETFOCUS('INETLAST'))

INETUPS->(DBGOTOP())

//SELECT INETUPS
//SEEK DATE()

//IF !FOUND()
 //	SEEK DATE()-1
//ENDIF
aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

@ 1,1 DCSAY 'Converted to Ups' SAYCOLOR GRA_CLR_WHITE,GRA_CLR_DARKBLUE

@ 1,40 DCSAY 'Converted AND Sold' SAYCOLOR GRA_CLR_BLACK,GRA_CLR_GREEN

@ 1,70 DCSAY '** Sortable Columns Right Click to Sort' SAYCOLOR GRA_CLR_BLACK,GRA_CLR_CYAN

@ 2,0 DCSAY 'Enter Search Key For Order Chosen' get cSearchkey GETID 'SEARCH' PICTURE '@!' ;
		      KEYBLOCK {|a,b,o|DC_BrowseAutoSeek(a,o,oBrowse)}

@ 2,60 DCSAY 'Sort Order ' get cIndexorder EDITPROTECT{|| TRUE }


@ 3.5,1 DCTOOLBAR oToolBar                               ;
	SIZE 140, 1.5

	DCADDBUTTON CAPTION 'Statistical Reports' SIZE 20,1.5;
		ACTION {|| _INETUPSTATS()} PARENT oToolbar

	DCADDBUTTON CAPTION 'Create Up Record' SIZE 20,1.5;
		ACTION {|| oRecord:=INETUPS->(DC_DBRECORD() : NEW()),;
				 INETUPS->(DC_DBSCATTER(oRecord)),;
		       _INET2UPREC(oRecord)} PARENT oToolbar

	DCADDBUTTON CAPTION 'Send eMail '                              ;
	SIZE 15                                              ;
	ACTION {||iNetUpemail()};
	PARENT oToolBar

	DCADDBUTTON CAPTION 'Review eMails '                              ;
		SIZE 15                                              ;
		ACTION {|| BROWEMAILLOG('',INETUPS->EMAIL)} ;
		PARENT oToolBar

   DCADDBUTTON CAPTION 'Refresh Vin Solutions '                              ;
		SIZE 20                                              ;
		ACTION {|| getvinsol(),INETUPS->(ORDSETFOCUS('INETLAST')),;
      INETUPS->(DBGOTOP()),   ;
      oBrowse:refreshall(),dc_getrefresh(getlist),obrowse:forcestable()} ;
		PARENT oToolBar


/* ----- Create browse ----- */

@ 6,1 DCBROWSE oBrowse ALIAS 'INETUPS'                 ;
	SIZE 140,20                                     ;
	PRESENTATION aPres;
	FREEZELEFT {1,2};
	SORTSCOLOR 0,2 ;
	SORTUCOLOR 1,6   ;
	COLOR {|| _INETUPCOLOR()}   ;
	EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITACROSS
	//DATALINK{||nREC:=INETUPS->(RECNO()) }

DCBROWSECOL FIELD INETUPS->DATERECVD ;
	HEADER "**DATERECVD" PARENT oBrowse    ;
	SORT{|| cIndexorder:='Date Received Order   ',INETUPS->(ORDSETFOCUS('INETDATE')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(DC_GETOBJECT(GETLIST,'SEARCH'))};
	EDITPROTECT {|| TRUE } ;
	WIDTH 5          ;

DCBROWSECOL FIELD INETUPS->LASTNAME ;
	HEADER "**Lastname" HCOLOR 1,5 PARENT oBrowse;
	WIDTH 10   ;
	SORT{|| cIndexorder:='Last Name Order   ',INETUPS->(ORDSETFOCUS('INETLAST')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(DC_GETOBJECT(GETLIST,'SEARCH'))}

DCBROWSECOL FIELD INETUPS->FIRSTNAME ;
	HEADER "First" HCOLOR 1,5 PARENT oBrowse    ;
	WIDTH 6
DCBROWSECOL FIELD INETUPS->FULLNAME ;
	HEADER "Full Name" HCOLOR 1,5 PARENT oBrowse    ;
	WIDTH 10


DCBROWSECOL FIELD INETUPS->EMAIL ;
	HEADER "**EMail" HCOLOR 1,5 PARENT oBrowse ;
	WIDTH 20  ;
	SORT{|| cIndexorder:='eMail Address Order   ',INETUPS->(ORDSETFOCUS('INETUPS')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(DC_GETOBJECT(GETLIST,'SEARCH'))}

DCBROWSECOL FIELD INETUPS->DAYPHONE ;
	HEADER "**Main Phone" HCOLOR 1,5 PARENT oBrowse       ;
	WIDTH 10 ;
	SORT{|| cIndexorder:='Day Phone Order   ',INETUPS->(ORDSETFOCUS('INETPHON')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(DC_GETOBJECT(GETLIST,'SEARCH'))}


DCBROWSECOL FIELD INETUPS->EVEPHON ;
	HEADER "**Alt Phone" HCOLOR 1,5 PARENT oBrowse        ;
	WIDTH 10 ;
	SORT{|| cIndexorder:='Evening Phone Order   ',INETUPS->(ORDSETFOCUS('INETEVE')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(DC_GETOBJECT(GETLIST,'SEARCH'))}

DCBROWSECOL FIELD INETUPS->cell ;
	HEADER "**Cell Phone" HCOLOR 1,5 PARENT oBrowse        ;
	WIDTH 10 ;
	SORT{|| cIndexorder:='Cell Phone Order      ',INETUPS->(ORDSETFOCUS('INETCELL')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(DC_GETOBJECT(GETLIST,'SEARCH'))}



DCBROWSECOL FIELD INETUPS->SOURCE ;
	HEADER "**Source" HCOLOR 1,5  PARENT oBrowse ;
	SORT{|| cIndexorder:='Source Order        ',INETUPS->(ORDSETFOCUS('INETSRCE')),dbgotop(),oBrowse:REFRESHALL(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(DC_GETOBJECT(GETLIST,'SEARCH'))};
	WIDTH 6

DCBROWSECOL FIELD INETUPS->TIMESTAMP ;
	HEADER "INETDATE" HCOLOR 1,5 PARENT oBrowse  ;
	EDITPROTECT {|| TRUE } ;
	WIDTH 8

DCBROWSECOL FIELD INETUPS->STREET ;
	HEADER "Street" HCOLOR 1,5 PARENT oBrowse   ;
	WIDTH 8
DCBROWSECOL FIELD INETUPS->CITY ;
	HEADER "City" HCOLOR 1,5 PARENT oBrowse  ;
	WIDTH 8

DCBROWSECOL FIELD INETUPS->STATE ;
	HEADER "State" HCOLOR 1,5 PARENT oBrowse      ;
	WIDTH 2

DCBROWSECOL FIELD INETUPS->ZIPCODE ;
	HEADER "**Zip" HCOLOR 1,5 PARENT oBrowse      ;
	WIDTH 5;
	SORT{|| cIndexorder:='Zipcode Order         ',INETUPS->(ORDSETFOCUS('INETZIP')),dbgotop(),oBrowse:REFRESHALL(),SETAPPFOCUS(DC_GETOBJECT(GETLIST,'SEARCH'))}

DCBROWSECOL FIELD INETUPS->CONLAST ;
	HEADER "SP Name " HCOLOR 1,5 PARENT oBrowse    ;
	WIDTH 6
DCBROWSECOL FIELD INETUPS->SM ;
	HEADER "SP Initials" HCOLOR 1,5 PARENT oBrowse    ;
	WIDTH 6


DCBROWSECOL FIELD INETUPS->COMMENT ;
	HEADER "Comment" HCOLOR 1,5 PARENT oBrowse      ;
	WIDTH 20



DCREAD GUI ;
	FIT ;
	title 'InterNet Lead Feed Browser';
	OPTIONS GETOPTIONS;
	EVAL {|| SETAPPFOCUS(DC_GETOBJECT(GETLIST,'SEARCH'))};
	BUTTONS DCGUI_BUTTON_EXIT
ReTURN nREC


static function _INETUPCOLOR()
	IF INETUPS->OPTOUT='CONVERTED'
		RETURN {GRA_CLR_WHITE,GRA_CLR_DARKBLUE}
	ENDIF
	IF INETUPS->OPTOUT='SOLD'
		RETURN {GRA_CLR_BLACK,GRA_CLR_GREEN}
	ENDIF
	RETURN {GRA_CLR_BLACK,GRA_CLR_WHITE}

static function _INET2UPREC(oRecord)
	LOCAL GETLIST:={},GETOPTIONS,xStatus,oUpRecord
	local cPhone:='',cArea:='',cUpid:='', nRec:=0 ,cComment:='',cSMName:=''
	DCGETOPTIONS SAYWIDTH 0 SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE

	/// make new up record object
	oUpRecord:=NEWUPS->(DC_DBRECORD(): NEW())
	NEWUPS->(DB_INIT(oUpRecord))
	/// populate it
	oUpRecord:WORKPHN:=oRecord:EVEPHON
	oUpRecord:CELLPHONE:=oRecord:CELL
	oUpRecord:COMMENTS:=oRecord:COMMENT
	oUpRecord:ORIGSOURCE:='I'

	/// make sense out of phone numbers
	IF !EMPTY(INETUPS->DAYPHONE)
		cPhone:=strtran(INETUPS->DAYPHONE,'-','')
		cPhone:=strtran(cPhone,'(','')
		cPhone:=strtran(cPhone,')','')
		cPhone:=alltrim(cPhone)
		cArea:=substr(cPhone,1,3)
		cUpId:=substr(cPhone,4,7)
		oUpRecord:UPID:=cUpid
		oUpRecord:AREA:=cArea
	  ELSE
		cPhone:=strtran(INETUPS->EVEPHON,'-','')
		cPhone:=strtran(cPhone,'(','')
		cPhone:=strtran(cPhone,')','')
		cPhone:=alltrim(cPhone)
		cArea:=substr(cPhone,1,3)
		cUpId:=substr(cPhone,4,7)
		oUpRecord:UPID:=cUpid
		oUpRecord:AREA:=cArea
	ENDIF
	IF EMPTY(cUpid)
		cPhone:=strtran(INETUPS->CELL,'-','')
		cPhone:=strtran(cPhone,'(','')
		cPhone:=strtran(cPhone,')','')
		cPhone:=alltrim(cPhone)
		cArea:=substr(cPhone,1,3)
		cUpId:=substr(cPhone,4,7)
		oUpRecord:UPID:=cUpid
		oUpRecord:AREA:=cArea
	ENDIF
	oUpRecord:EMAIL:=INETUPS->EMAIL
	oUpRecord:FIRSTNAME:=oRecord:FIRSTNAME
	oUpRecord:LASTNAME:=oRecord:LASTNAME
	oUpRecord:DATEIN :=DATE()
	oUpRecord:STREET :=oRecord:STREET
	oUpRecord:CITYSTATE :=oRecord:CITY
	oUpRecord:NOTES :=oRecord:COMMENT
	oUpRecord:UPSTATUS :='A'
	oUpRecord:BEBACK3 :=Date()
	IF empty(oUpRecord:UPID)
		oUprecord:UPID:=space(7)
	ENDIF

	@ 5,0 DCSAY 'This selection will assign this prospect to a Sales Person and create an Up Record.'
	@ 6,0 DCSAY 'You Must enter some tracking ID such as a phone number ' get oUpRecord:UPid valid{|| IIF(EMPTY(oUpRecord:UPID),.f.,.t.)}
	@ 8,0 DCSAY 'Please enter the SalesPersons Initial ID assigned to this prospect' get  oUpRecord:SALESMAN valid{|| _IsSMGood(oUpRecord:SALESMAN,@cSMName) } PICTURE '@!'
	DCREAD GUI to xStatus BUTTONS 2 ENTEREXIT OPTIONS GETOPTIONS FIT TITLE  'Assign SalesPerson'
	IF !xStatus
		RETURN NIL
	ENDIF

	/// NOW CREAT RECORD
	SELECT NEWUPS
	IF !EMPTY(cUpid)
	 ordsetfocus('NEWUPIN')
	 SEEK cUpid
	 IF found()
		dc_winalert('This Up already exists')
		INETUPS->(ORDSETFOCUS('INETLAST'))
      INETUPS->(DBGOTOP())
		return nil
	 ENDIF
	ENDIF
	IF !empty(oUpRecord:EMAIL)
		ORDSETFOCUS('NEWUPEML')
		SEEK oUpRecord:EMAIL
		IF found()
		 dc_winalert('This Up already exists')
		 return nil
	   ENDIF
	ENDIF


	NEWUPS->(DC_DBGATHER(oUpRecord,.t.))


	DC_WINALERT('UP ADDED TO SYSTEM')


	SELECT INETUPS
	IF INETUPS->(DC_RECLOCK())
		nRec:=INETUPS->(RECNO())
		REPLACE INETUPS->OPTOUT WITH 'CONVERTED'
		REPLACE INETUPS->SM WITH oUpRecord:Salesman
		REPLACE INETUPS->CONLAST WITH cSMName
		INETUPS->(DBRUNLOCK(nRec))
	ENDIF

	///now run main up entry edit screen
	UPQUOTE(NEWUPS->(RECNO()))

RETURN NIL




Static Function _IsSMGood(cSM,cName)
	IF EMPTY(cSM)
		RETURN FALSE
	ENDIF
	select SM
	SM->(DBSEEK(cSM))
	IF found()
		cName:=SM->Name
		RETURN TRUE
	ENDIF
	RETURN FALSE



STATIC FUNCTION _INETUPSTATS()
	LOCAL GETLIST:={} , GETOPTIONS , oSortgroup, nSort:=1  , dDateFrom:=ctod('  /  /  '), dDateTo:=date() , lJustSold:=.f. , xstatus
	local nSold:=0


	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE

	@ 1,0 DCGROUP oSortgroup SIZE 20,10 Caption 'Sort Choices'
	@ 1,1 DCRADIO nSort VALUE 1 PROMPT 'By Source  ' COLOR 1,5 ;
      PARENT oSortgroup
	@ 3,1 DCRADIO nSort VALUE 2 PROMPT 'By Sales Person  ' COLOR 1,4 ;
		 PARENT oSortgroup
	@ 5,1 DCRADIO nSort VALUE 3 PROMPT 'By Zip Code  ' COLOR 1,3 ;
		 PARENT oSortgroup
	@ 7,1 DCRADIO nSort VALUE 4 PROMPT 'By Last Name  ' COLOR 1,6 ;
		 PARENT oSortgroup
	@ 9,1 DCRADIO nSort VALUE 5 PROMPT 'By eMail  ' COLOR 1,7 ;
		 PARENT oSortgroup
	@ 14,0 DCSAY 'Enter Date Range From ' get dDateFrom
	@ 15,0 DCSAY 'Enter Date Range To   ' get dDateto

	//@ 18,0 DCCHECKBOX lJustSold PROMPT 'Check here to show ONLY SOLD prospects. '

	DCREAD GUI MODAL to xStatus ADDBUTTONS OPTIONS GETOPTIONS FIT TITLE 'InterNet Up Stats' EVAL{|o| SETAPPWINDOW(o)}
	IF !xStatus
		return nil

	ENDIF

	SELECT INETUPS
	IF nSort=1
		ORDSETFOCUS('INETSRCE')
		INETUPS->(DBGOTOP())
		_PRINTINETBYSRCE(dDatefrom,ddateto, 'Ups by Internet Source')
		return nil
	ENDIF
	IF nSort=2
		ORDSETFOCUS('INETSM')
		INETUPS->(DBGOTOP())
		_PRINTINETBYSM(dDatefrom,ddateto, 'InterNet Ups by Sales Person')
		return nil
	ENDIF
	IF nSort=3
		ORDSETFOCUS('INETZIP')
		INETUPS->(DBGOTOP())
		_PRINTINETBYZIP(dDatefrom,ddateto, 'InterNet Ups by Zipcode')
		return nil
	ENDIF
	IF nSort=4
		ORDSETFOCUS('INETLAST')
		INETUPS->(DBGOTOP())
		_PRINTINETBYLAST(dDatefrom,ddateto, 'InterNet Ups by Last Name')
		return nil
	ENDIF
	IF nSort=5
		ORDSETFOCUS('INETUPS')
		INETUPS->(DBGOTOP())
		_PRINTINETBYEMAIL(dDatefrom,ddateto, 'InterNet Ups by EMAIL Address')
		return nil
	ENDIF

	return nil






  STATIC FUNCTION _PRINTINETBYSRCE(dDatefrom,dDateto,cMess)
	LOCAL nOrient:=2, nDefmode:=4,nCopies:=1, cOutfile:='',cSentfont:='8.Courier New', nPage:=0 , oPrinter
	LOCAL cSource:=space(10)    , nRecCount:=0 , nTotCount:=1 , nSoldCount:=0 ,nTotSoldCount:=0 ,nConverted:=0 ,nTotConverted:=0
	TOP_MAR(3)
	BOT_MAR(2)
	PAGE_LEN(62)
	IF !Print_Choice( 'iNetReports', @nCopies,' ',.F.,@nOrient,@cSentfont,@nDefMode,cOutfile )
			 RETURN NIL
	ENDIF

	IF nOrient=1
			PAGE_LEN(80)
	ENDIF

	IF nDefMode=8
		PAGE_LEN(32000)
	ENDIF


	IF !PrintOn('iNetReports', oPrinter, nOrient, cSentfont, nCopies)
		  RETURN NIL
	ENDIF

	_INETUPSHDR(@nPage,dDatefrom,dDateto, cMess)

	DO WHILE !(INETUPS->(EOF()))
			  cSource:=INETUPS->SOURCE
			  @ DC_PRINTERROW()+1,0 DCPRINT SAY cSource
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF

			  nReccount:=0
			  nSoldCount:=0
			  nConverted:=0
			  DO WHILE INETUPS->SOURCE=cSource
				IF !deleted()
					IF INETUPS->DATERECVD >=dDatefrom
						IF INETUPS->DATERECVD <=dDateto
							nReccount++
							nTotCount++
							_printUpsLine(@nSoldcount,@nTotSoldCount,@nConverted,@nTotConverted)
						ENDIF
						IF dcpageeject()
							 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
						ENDIF
					ENDIF
				endif
				skip alias INETUPS
			  ENDDO
			  SKIPALINE()
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF
			  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Record Count '+ str(nReccount) FONT '10.Courier New Bold'
			  SKIPALINE()
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF
			  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Converted Count '+ str(nConverted) FONT '10.Courier New Bold'

			  SKIPALINE()
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF
			  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Sold Count '+ str(nSoldCount) FONT '10.Courier New Bold'

			  SKIPALINE()
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF

	  ENDDO
	  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Records Counted.. ' + str(nTotCount)
	  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Converted ' + str(nTotConverted)


	  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Sold.. ' + str(nTotSoldCount)



	  PRINTOFF(oPrinter)
	  return nil

  STATIC FUNCTION _PRINTINETBYSM(dDatefrom,dDateto,cMess)
	LOCAL nOrient:=2, nDefmode:=4,nCopies:=1, cOutfile:='',cSentfont:='8.Courier New', nPage:=0 , oPrinter
	LOCAL cSM:=space(2)    , nRecCount:=0 , nTotCount:=1, nSoldcount:=0 ,nTotSoldCount:=0 ,nConverted:=0 ,nTotConverted:=0


	TOP_MAR(3)
	BOT_MAR(2)
	PAGE_LEN(62)
	IF !Print_Choice( 'iNetReports', @nCopies,' ',.F.,@nOrient,@cSentfont,@nDefMode,cOutfile )
			 RETURN NIL
	ENDIF

	IF nOrient=1
			PAGE_LEN(80)
	ENDIF

	IF nDefMode=8
		PAGE_LEN(32000)
	ENDIF


	IF !PrintOn('iNetReports', oPrinter, nOrient, cSentfont, nCopies)
		  RETURN NIL
	ENDIF

	_INETUPSHDR(@nPage,dDatefrom,dDateto, cMess)

	DO WHILE !(INETUPS->(EOF()))
			  cSm:=INETUPS->SM
			  nReccount:=0
			  nSoldCount:=0
			  DO WHILE INETUPS->SM=cSM
				IF INETUPS->DATERECVD >=dDatefrom
					IF INETUPS->DATERECVD <=dDateto
						nReccount++
						nTotCount++
						_printUpsLine(@nSoldcount,@nTotSoldCount,@nConverted,@nTotConverted)
					ENDIF
					IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
					ENDIF
				ENDIF
				skip alias INETUPS
			  ENDDO
			  SKIPALINE()
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF
			  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Record Count '+ str(nReccount) FONT '10.Courier New Bold'
			  SKIPALINE()
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF
			  SKIPALINE()
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF
			  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Converted Count '+ str(nConverted) FONT '10.Courier New Bold'

			  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Sold.. ' + str(nSoldCount)

			  SKIPALINE()
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF

	  ENDDO
	  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Records Counted.. ' + str(nTotCount)
	  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Converted ' + str(nTotConverted)


	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Sold.. ' + str(nTotSoldCount)
	  PRINTOFF(oPrinter)
	  return nil

  STATIC FUNCTION _PRINTINETBYZIP(dDatefrom,dDateto,cMess)
	LOCAL nOrient:=2, nDefmode:=4,nCopies:=1, cOutfile:='',cSentfont:='8.Courier New', nPage:=0 , oPrinter
	LOCAL cZip:=space(5)    , nRecCount:=0 , nTotCount:=1 ,nSoldcount:=0 ,nTotSoldCount:=0,nConverted:=0 ,nTotConverted:=0


	TOP_MAR(3)
	BOT_MAR(2)
	PAGE_LEN(62)
	IF !Print_Choice( 'iNetReports', @nCopies,' ',.F.,@nOrient,@cSentfont,@nDefMode,cOutfile )
			 RETURN NIL
	ENDIF

	IF nOrient=1
			PAGE_LEN(80)
	ENDIF

	IF nDefMode=8
		PAGE_LEN(32000)
	ENDIF


	IF !PrintOn('iNetReports', oPrinter, nOrient, cSentfont, nCopies)
		  RETURN NIL
	ENDIF

	_INETUPSHDR(@nPage,dDatefrom,dDateto, cMess)

	DO WHILE !(INETUPS->(EOF()))
			  cZip:=INETUPS->zipcode
			  nReccount:=0
			  nSoldCount:=0
			  DO WHILE INETUPS->ZIPCODE=cZip
				IF INETUPS->DATERECVD >=dDatefrom
					IF INETUPS->DATERECVD <=dDateto
						nReccount++
						nTotCount++
						_printUpsLine(@nSoldcount,@nTotSoldCount,@nConverted,@nTotConverted)
					ENDIF
					IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
					ENDIF
				ENDIF
				skip alias INETUPS
			  ENDDO
			  SKIPALINE()
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF
			  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Source Record Count '+ str(nReccount) FONT '10.Courier New Bold'
			  SKIPALINE()
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF
			  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Converted Count '+ str(nConverted) FONT '10.Courier New Bold'


			  SKIPALINE()
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF
			  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Sold.. ' + str(nSoldCount)
			  SKIPALINE()
			  IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
			  ENDIF

	  ENDDO
	  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Records Counted.. ' + str(nTotCount)
	  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Converted ' + str(nTotConverted)
	  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Sold.. ' + str(nTotSoldCount)


	  PRINTOFF(oPrinter)
	  return nil

  STATIC FUNCTION _PRINTINETBYLAST(dDatefrom,dDateto,cMess)
	LOCAL nOrient:=2, nDefmode:=4,nCopies:=1, cOutfile:='',cSentfont:='8.Courier New', nPage:=0 , oPrinter
	LOCAL nTotCount:=1 ,nSoldcount :=0   ,nTotSoldCount:=0 ,nConverted:=0 ,nTotConverted:=0,nReccount:=0


	TOP_MAR(3)
	BOT_MAR(2)
	PAGE_LEN(62)
	IF !Print_Choice( 'iNetReports', @nCopies,' ',.F.,@nOrient,@cSentfont,@nDefMode,cOutfile )
			 RETURN NIL
	ENDIF

	IF nOrient=1
			PAGE_LEN(80)
	ENDIF

	IF nDefMode=8
		PAGE_LEN(32000)
	ENDIF


	IF !PrintOn('iNetReports', oPrinter, nOrient, cSentfont, nCopies)
		  RETURN NIL
	ENDIF

	_INETUPSHDR(@nPage,dDatefrom,dDateto, cMess)

	DO WHILE !(INETUPS->(EOF()))
			  nReccount:=0
				IF INETUPS->DATERECVD >=dDatefrom
					IF INETUPS->DATERECVD <=dDateto
						nTotCount++
						_printUpsLine(@nSoldcount,@nTotSoldCount,@nConverted,@nTotConverted)
					ENDIF
					IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
					ENDIF
				ENDIF
				SKIP ALIAS INETUPS
  ENDDO
  SKIPALINE()
  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
  ENDIF
  SKIPALINE()
  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
  ENDIF
  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Records Counted.. ' + str(nTotCount)
  SKIPALINE()
  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
  ENDIF
  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Converted ' + str(nTotConverted)
  SKIPALINE()
  IF dcpageeject()
  	 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
  ENDIF
  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Sold.. ' + str(nSoldCount)

  PRINTOFF(oPrinter)
  return nil

  STATIC FUNCTION _PRINTINETBYEMAIL(dDatefrom,dDateto,cMess)
	LOCAL nOrient:=2, nDefmode:=4,nCopies:=1, cOutfile:='',cSentfont:='8.Courier New', nPage:=0 , oPrinter
	LOCAL nTotCount:=1 ,nSoldcount:=0 ,nTotSoldCount:=0 ,nConverted:=0 ,nTotConverted:=0,nReccount:=0


	TOP_MAR(3)
	BOT_MAR(2)
	PAGE_LEN(62)
	IF !Print_Choice( 'iNetReports', @nCopies,' ',.F.,@nOrient,@cSentfont,@nDefMode,cOutfile )
			 RETURN NIL
	ENDIF

	IF nOrient=1
			PAGE_LEN(80)
	ENDIF

	IF nDefMode=8
		PAGE_LEN(32000)
	ENDIF


	IF !PrintOn('iNetReports', oPrinter, nOrient, cSentfont, nCopies)
		  RETURN NIL
	ENDIF

	_INETUPSHDR(@nPage,dDatefrom,dDateto, cMess)

	DO WHILE !(INETUPS->(EOF()))
			  nReccount:=0
				IF INETUPS->DATERECVD >=dDatefrom
					IF INETUPS->DATERECVD <=dDateto
						nTotCount++
						_printUpsLine(@nSoldcount,@nTotSoldCount,@nConverted,@nTotConverted)
					ENDIF
					IF dcpageeject()
						 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
					ENDIF
				ENDIF
				SKIP ALIAS INETUPS
  ENDDO
  SKIPALINE()
  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
  ENDIF
  SKIPALINE()
  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
  ENDIF
  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Records Counted.. ' + str(nTotCount)
  SKIPALINE()
  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
  ENDIF
  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Converted ' + str(nTotConverted)

  SKIPALINE()
	  IF dcpageeject()
	  		 _INETUPSHDR(@nPage,dDatefrom,dDateto,cMess)
	  ENDIF
	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Sold.. ' + str(nSoldCount)


	  PRINTOFF(oPrinter)
	  return nil


	static function _printUpsline(nSoldcount,nTotSoldCount,nConverted,nTotConverted)
		local nRec:=0,cMess:=''
		@ DC_PRINTERROW()+1,0 DCPRINT SAY SUBSTR(INETUPS->LASTNAME,1,15)
		@ DC_PRINTERROW(),16 DCPRINT SAY SUBSTR(INETUPS->FIRSTNAME,1,10)
		@ DC_PRINTERROW(),26 DCPRINT SAY INETUPS->DATERECVD
		@ DC_PRINTERROW(),35 DCPRINT SAY INETUPS->SOURCE
		@ DC_PRINTERROW(),48 DCPRINT SAY INETUPS->CONLAST
		@ DC_PRINTERROW(),60 DCPRINT SAY SUBSTR(INETUPS->EMAIL,1,30)
		@ DC_PRINTERROW(),92 DCPRINT SAY INETUPS->DAYPHONE
		@ DC_PRINTERROW(),107 DCPRINT SAY INETUPS->ZIPCODE
		@ DC_PRINTERROW(),115 DCPRINT SAY SUBSTR(INETUPS->CITY,1,10)+','+SUBSTR(INETUPS->STATE,1,3)
		SELECT NEWUPS
		ORDSETFOCUS('NEWUPEML')
		NEWUPS->(DBGOTOP())
		SEEK INETUPS->EMAIL
		IF FOUND()
			cMess:='CONVERTED'
			nConverted++
			nTotConverted++
			IF NEWUPS->UPSTATUS='S'
				nSoldcount++
				nTotSoldCount++
				cMess:='SOLD'
			ENDIF
		ELSE
			SELECT NEWUPS
			ORDSETFOCUS('NEWUPNM')
		   NEWUPS->(DBGOTOP())
		   SEEK INETUPS->LASTNAME
		   IF FOUND()
			 IF UPPER(SUBSTR(NEWUPS->FIRSTNAME,1,2)) = UPPER(SUBSTR(INETUPS->FIRSTNAME,1,2))
			  cMess:='CONVERTED'
			  nConverted++
			  nTotConverted++
			  IF NEWUPS->UPSTATUS='S'
				cMess:='SOLD'
				nSoldcount++
				nTotSoldCount++
			  ENDIF

			  SELECT INETUPS
			  IF INETUPS->(DC_RECLOCK())
				nRec:=inetups->(recno())
				REPLACE INETUPS->OPTOUT WITH cMess
				DBRUNLOCK(nRec)
			  ENDIF
			 ENDIF
			ENDIF
		ENDIF
		@ DC_PRINTERROW(),132 DCPRINT SAY cMess FONT '10.Courier New Bold'

		RETURN NIL


	static function _INETUPSHDR(nPage,dDatefrom,dDateto,cMess)
		nPage++
		@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Internet Ups Reports  '+ dtoc(dDateFrom)+'    to   '+ dtoc(dDateto)+'  '+cMess +'  '+ALLTRIM(STR(nPage))
		@ DC_PRINTERROW()+1,0 DCPRINT SAY 'LASTNAME'
		@ DC_PRINTERROW(),16 DCPRINT SAY 'FIRSTNAME'
		@ DC_PRINTERROW(),26 DCPRINT SAY 'DATE'
		@ DC_PRINTERROW(),35 DCPRINT SAY 'SOURCE'
		@ DC_PRINTERROW(),48 DCPRINT SAY 'SALESPERSON'
		@ DC_PRINTERROW(),62 DCPRINT SAY 'EMAIL'
		@ DC_PRINTERROW(),92 DCPRINT SAY 'DAYPHONE'
		@ DC_PRINTERROW(),107 DCPRINT SAY 'ZIPCODE'
		@ DC_PRINTERROW(),116 DCPRINT SAY 'CITY/ST'



		RETURN NIL








FUNCTION BROWBYDT()
LOCAL GETLIST:={} , xStatus , cProgram , GETOPTIONS ,oRecord
LOCAL dDatex:=CTOD('  /  /  ') ,oToolbar , oBrowse ,aPres , lNoShowDead:=.t.
DCGETOPTIONS ;
SAYWIDTH 0 SAYFONT '10.COURIER NEW' GETFONT '10.COURIER NEW';
AUTORESIZE

cProgram:='BROWBYDT'

IF.NOT.SECURITY(@cProgram)
	DC_WINALERT('SECURITY VIOLATION')
 RETURN NIL
ENDIF

IF !OPENUPSFILES()
	CLOSE ALL
   RETURN NIL
ENDIF

NEWUPS->(ORDSETFOCUS('NEWUPDT'))


XSTATUS:=.T.
dDatex:=(DATE()-1)


@ 12,0 DCSAY 'Enter The Date To Go Back To Begin Browse' GET dDatex SAYCOLOR 9,0 SAYFONT '10.COURIER' GETFONT '12.COURIER'
//@ 14,0 DCCHECKBOX lNoShowDead PROMPT 'Click here to NOT show deals already DEAD'
DCREAD GUI TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'ENTER DATE TO BEGIN BROWSE'
IF !XSTATUS
 RETURN NIL
ENDIF
SET SOFTSEEK ON
SELECT NEWUPS
GOTO TOP
SEEK dDatex
SET SOFTSEEK OFF
IF lNoShowDead
	SET FILTER TO NEWUPS->upstatus <> 'D'
ENDIF
aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

@ 1,1 DCTOOLBAR oToolBar                               ;
	SIZE 140, 1.5 BUTTONSIZE 15,1.5

DCADDBUTTON CAPTION 'Desking Tool;Quote '                              ;
	;//ACTION {||oRecord:=NEWUPS->(DC_DBRECORD():NEW()),;
	;//NEWUPS->(DC_DBSCATTER(oRecord)),;
	ACTION {|| UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)};
	PARENT oToolBar


DCADDBUTTON CAPTION '2.Browse New;Vehicles  '                              ;
	ACCELKEY {ASC("2")};
	ACTION {||SBROWBYDESC(),DC_GETREFRESH(GETLIST)  };
	PARENT oToolBar

DCADDBUTTON CAPTION '3.Browse Used;Vehicles '                              ;
	ACCELKEY {ASC("3")};
	ACTION {||SBROWBYDESC(),DC_GETREFRESH(GETLIST)  };
	PARENT oToolBar

DCADDBUTTON CAPTION '4.Maintain Up;Desking Tool '                              ;
	ACCELKEY {ASC("4")};
	ACTION {|| UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)  };
	PARENT oToolBar

DCADDBUTTON CAPTION '5.View Up;Comments '                              ;
	ACCELKEY {ASC("5")};
	ACTION {|| UPNOTES(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(obROWSE:GETCOLUMN(1))  };
	PARENT oToolBar
DCADDBUTTON CAPTION '6.Print UPSheet '                              ;
	ACCELKEY {ASC("6")};
	ACTION {||PRINTUPSHEET(NEWUPS->(RECNO()))};
	PARENT oToolBar

DCADDBUTTON CAPTION '7.Send eMail '                              ;
	ACCELKEY {ASC("7")};
	ACTION {||upemail()};
	PARENT oToolBar


DCADDBUTTON CAPTION '8.Appointment; History '                              ;
	ACCELKEY {ASC("8")};
	ACTION {||_chkapptsbrowse(NEWUPS->UpId)};
	PARENT oToolBar

DCADDBUTTON CAPTION '9.Review eMails '                              ;
		SIZE 15                                              ;
		ACCELKEY {ASC("9")};
		ACTION {|| BROWEMAILLOG('',NEWUPS->EMAIL)} ;
		PARENT oToolBar


/* ----- Create browse ----- */


@ 3,0 DCSAY 'COLOR CODES'
@ 3,15 DCSAY 'Sold' SAYCOLOR 1,5
@ 3,30 DCSAY 'Active' SAYCOLOR 1,0
@ 3,45 DCSAY 'Tickle' SAYCOLOR 1,15
@ 3,60 DCSAY 'DEAD' SAYCOLOR 0,1
@ 3,75 DCSAY 'Rescue' SAYCOLOR 2,3
@ 3,90 DCSAY 'PhoneUp' SAYCOLOR 12,7
@ 3,105 DCSAY 'InterNetUp' SAYCOLOR 1,4


@ 5,1 DCBROWSE oBrowse ALIAS 'NEWUPS'                 ;
	SIZE 140,20                                     ;
	PRESENTATION aPres;
	FREEZELEFT {1,2,3} ;
	color{|| getupcolor(NEWUPS->UPSTATUS)};
	EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITACROSS ;
	//DATALINK {||MAINUP(.T.)}

DCBROWSECOL FIELD NEWUPS->Upstatus ;
	HEADER "Status" PARENT oBrowse  ;
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1

DCBROWSECOL FIELD NEWUPS->ORIGSOURCE ;
	HEADER "Source" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1


DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "UPID/PHN#" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 7

DCBROWSECOL FIELD NEWUPS->PREFMETHOD ;
	HEADER "PREFMETHOD" PARENT oBrowse  ;
	WIDTH 4



DCBROWSECOL FIELD NEWUPS->SALESMAN ;
	HEADER "SLSPRS" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->LASTNAME                     ;
	HEADER "LASTNAME" HCOLOR 9,3 PARENT oBrowse     ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->FIRSTNAME ;
	HEADER "FIRSTNAME" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->DATEIN ;
	HEADER "LastDateIn" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->beback1 ;
	HEADER "Beback1" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->beback2 ;
	HEADER "Beback2" PARENT oBrowse  ;
	WIDTH 8

DCBROWSECOL FIELD NEWUPS->beback3 ;
	HEADER "OrigDateIn" PARENT oBrowse  ;
	WIDTH 8

DCBROWSECOL FIELD NEWUPS->EMAIL ;
	HEADER "EMAIL" PARENT oBrowse  ;
	WIDTH 20

DCBROWSECOL FIELD NEWUPS->STREET ;
	HEADER "STREET" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->CITYSTATE ;
	HEADER "CITYSTATE" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->ZIPCODE ;
	HEADER "ZIPCODE" PARENT oBrowse  ;
	WIDTH 5
DCBROWSECOL FIELD NEWUPS->AREA ;
	HEADER "AREACODE" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "PHONE" PARENT oBrowse  ;
	WIDTH 7
DCBROWSECOL FIELD NEWUPS->WORKPHN ;
	HEADER "WORKPHN" PARENT oBrowse  ;
	WIDTH 12
DCBROWSECOL FIELD NEWUPS->cellphone ;
	HEADER "CELLPHN" PARENT oBrowse  ;
	WIDTH 10

DCBROWSECOL FIELD NEWUPS->NEWUSED ;
	HEADER "NEWUSED" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->INTEREST ;
	HEADER "INTEREST" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->SALECODE ;
	HEADER "SALECODE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->TOMGR ;
	HEADER "TOMGR" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->COMMENTS ;
	HEADER "COMMENT" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->COMMENT2 ;
	HEADER "COMMENT2" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->ADSOURCE ;
	HEADER "ADSOURCE" PARENT oBrowse  ;
	WIDTH 4
DCBROWSECOL FIELD NEWUPS->DOWNDEAL ;
	HEADER "DOWN" PARENT oBrowse  ;
	WIDTH 1
DCREAD GUI ;
	FIT ;
	EVAL {||SETAPPFOCUS(obROWSE:GETCOLUMN(1))};
	OPTIONS GETOPTIONS ;
	BUTTONS DCGUI_BUTTON_EXIT ;
	title 'Automan Up Browser by Date'
CLOSE ALL
ReTURN nil

FUNCTION BLOWBYBLOW()
LOCAL GETLIST:={} , xStatus , cProgram , GETOPTIONS ,cSm:=space(2)
LOCAL dDatex:=CTOD('  /  /  ') ,oToolbar , oBrowse ,aPres , lNoShowDead:=.t.
DCGETOPTIONS ;
SAYWIDTH 0 SAYFONT '10.COURIER NEW' GETFONT '10.COURIER NEW';
AUTORESIZE

cProgram:='BLOWBYBLOW'

IF.NOT.SECURITY(@cProgram)
	DC_WINALERT('SECURITY VIOLATION')
 RETURN NIL
ENDIF

IF !OPENUPSFILES()
	CLOSE ALL
   RETURN NIL
ENDIF

NEWUPS->(ORDSETFOCUS('NEWUPDT'))


XSTATUS:=.T.
dDatex:=(DATE()-1)
@ 12,0 DCSAY 'Enter The Date To Go Back To Begin Browse' GET dDatex SAYCOLOR 9,0 SAYFONT '10.COURIER' GETFONT '12.COURIER'
@ 13,0 DCSAY 'Enter the Sales Person to Browse for or leave blank for all.' get cSm
DCREAD GUI TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'ENTER DATE TO BEGIN BROWSE'
IF !XSTATUS
 RETURN NIL
ENDIF
SET SOFTSEEK ON
SELECT NEWUPS
GOTO TOP
SEEK dDatex
SET SOFTSEEK OFF
IF !empty(cSm)
	SET FILTER TO NEWUPS->SALESMAN = cSm
ENDIF

aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

@ 1,1 DCTOOLBAR oToolBar                               ;
	SIZE 130, 1.5 BUTTONSIZE 12,1.5


DCADDBUTTON CAPTION '1.Maintain Up;Record '                              ;
	ACCELKEY {ASC("1")};
	ACTION {|| UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)  };
	PARENT oToolBar

DCADDBUTTON CAPTION '2.View Up;Comments '                              ;
	ACCELKEY {ASC("2")};
	ACTION {|| UPNOTES(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(obROWSE:GETCOLUMN(1))  };
	PARENT oToolBar

DCADDBUTTON CAPTION '3.Send eMail '                              ;
	ACCELKEY {ASC("3")};
	ACTION {||upemail()};
	PARENT oToolBar

DCADDBUTTON CAPTION '4.Appointment; History '                              ;
	ACCELKEY {ASC("4")};
	ACTION {||_chkapptsbrowse(NEWUPS->UpId)};
	PARENT oToolBar

DCADDBUTTON CAPTION '5.Print/Output; Reports '                              ;
	ACCELKEY {ASC("4")};
	ACTION {||_blbyblreps(dDatex,cSm)};
	PARENT oToolBar



/* ----- Create browse ----- */

@ 2,0 DCSAY

@ 3,0 DCSAY 'SOURCE COLOR CODES'
@ 4,0 DCSAY 'Floor Traffic' SAYCOLOR 1,0
@ 4,15 DCSAY '   PhoneUp   ' SAYCOLOR 12,7
@ 4,30 DCSAY 'InterNetUp   ' SAYCOLOR 1,4
@ 3,51 DCSAY 'SHOWED COLOR CODES'
@ 4,51 DCSAY 'NoShow ' SAYCOLOR 1,3
@ 4,60 DCSAY 'Showed ' SAYCOLOR 1,5
@ 4,70 DCSAY 'Pending' SAYCOLOR 0,9
@ 4,80 DCSAY 'No Appt' SAYCOLOR 1,15

@ 6,1 DCBROWSE oBrowse ALIAS 'NEWUPS'                 ;
	SIZE 120,20                                     ;
	PRESENTATION aPres;
	FREEZELEFT {1,2,3} ;
	//color{|| getupcolor(NEWUPS->UPSTATUS)}
	///EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITACROSS


DCBROWSECOL FIELD NEWUPS->DATEIN ;
	HEADER "LastDateIn" PARENT oBrowse  ;
	WIDTH 5

DCBROWSECOL FIELD NEWUPS->SALESMAN ;
	HEADER "SLSPRS" PARENT oBrowse  ;
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->Upstatus ;
	HEADER "Status" PARENT oBrowse  ;
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1

DCBROWSECOL FIELD NEWUPS->ORIGSOURCE ;
	HEADER "Source" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1

DCBROWSECOL FIELD NEWUPS->LASTNAME                     ;
	HEADER "LASTNAME" HCOLOR 9,3 PARENT oBrowse     ;
	WIDTH 10

DCBROWSECOL FIELD NEWUPS->APPOINT ;
	HEADER "Appt" PARENT oBrowse COLOR {|| IIF(NEWUPS->APPOINT ,{GRA_CLR_WHITE,GRA_CLR_DARKGREEN} ,{GRA_CLR_WHITE,GRA_CLR_DARKRED} )} ;
	WIDTH 3

DCBROWSECOL DATA {|| _CHK4SHOW(NEWUPS->UPID)} ;
	HEADER 'Showed' PARENT oBrowse COLOR {|| _CHK4SHOWCLR(NEWUPS->UPID)};
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->DEMOED ;
	HEADER "Demo" PARENT oBrowse COLOR {|| IIF(NEWUPS->DEMOED ,{GRA_CLR_WHITE,GRA_CLR_DARKGREEN} ,{GRA_CLR_WHITE,GRA_CLR_DARKRED} )} ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->TOED ;
	HEADER "TOed" PARENT oBrowse COLOR {|| IIF(NEWUPS->TOED ,{GRA_CLR_WHITE,GRA_CLR_DARKGREEN} ,{GRA_CLR_WHITE,GRA_CLR_DARKRED} )} ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->WRITEUP ;
	HEADER "WritenUp" PARENT oBrowse COLOR {|| IIF(NEWUPS->WRITEUP ,{GRA_CLR_WHITE,GRA_CLR_DARKGREEN} ,{GRA_CLR_WHITE,GRA_CLR_DARKRED} )} ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->SOLD ;
	HEADER "Sold" PARENT oBrowse COLOR {|| IIF(NEWUPS->SOLD ,{GRA_CLR_BLACK,GRA_CLR_GREEN} ,{GRA_CLR_WHITE,GRA_CLR_DARKRED} )} ;
	WIDTH 3

DCBROWSECOL DATA {|| _xchkbillstat()} ;
	HEADER "Delivered" PARENT oBrowse COLOR {|| IIF(NEWUPS->BILLSTAT='D',{GRA_CLR_BLACK,GRA_CLR_GREEN} ,{GRA_CLR_BLACK,GRA_CLR_PALEGRAY} )}

DCBROWSECOL DATA {|| _chksendltr()}  ;
	HEADER "DeliveryFollowUp" PARENT oBrowse COLOR {|| IIF( NEWUPS->SENDLTR='Y',{GRA_CLR_BLACK,GRA_CLR_GREEN} ,{GRA_CLR_BLACK,GRA_CLR_PALEGRAY} )}


DCBROWSECOL FIELD NEWUPS->beback3 ;
	HEADER "OrigDateIn" PARENT oBrowse  ;
	WIDTH 5


DCBROWSECOL FIELD NEWUPS->beback1 ;
	HEADER "Beback1" PARENT oBrowse  ;
	WIDTH 5
DCBROWSECOL FIELD NEWUPS->beback2 ;
	HEADER "Beback2" PARENT oBrowse  ;
	WIDTH 5


DCBROWSECOL FIELD NEWUPS->EMAIL ;
	HEADER "EMAIL" PARENT oBrowse  ;
	WIDTH 20

DCBROWSECOL FIELD NEWUPS->AREA ;
	HEADER "AREACODE" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "PHONE" PARENT oBrowse  ;
	WIDTH 7
DCBROWSECOL FIELD NEWUPS->WORKPHN ;
	HEADER "WORKPHN" PARENT oBrowse  ;
	WIDTH 12
DCBROWSECOL FIELD NEWUPS->cellphone ;
	HEADER "CELLPHN" PARENT oBrowse  ;
	WIDTH 10

DCBROWSECOL FIELD NEWUPS->NEWUSED ;
	HEADER "NEWUSED" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->INTEREST ;
	HEADER "INTEREST" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->SALECODE ;
	HEADER "SALECODE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->TOMGR ;
	HEADER "TOMGR" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->COMMENTS ;
	HEADER "COMMENT" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->COMMENT2 ;
	HEADER "COMMENT2" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->ADSOURCE ;
	HEADER "ADSOURCE" PARENT oBrowse  ;
	WIDTH 4
DCBROWSECOL FIELD NEWUPS->DOWNDEAL ;
	HEADER "DOWN" PARENT oBrowse  ;
	WIDTH 1
DCREAD GUI ;
	FIT ;
	EVAL {||SETAPPFOCUS(obROWSE:GETCOLUMN(1))};
	OPTIONS GETOPTIONS ;
	BUTTONS DCGUI_BUTTON_EXIT ;
	title 'Automan Up Browser by Date'
CLOSE ALL
ReTURN nil


STATIC FUNCTION _xchkbillstat()

IF NEWUPS->BILLSTAT='D'
   return .T.
else
	return FALSE
endif
return FALSE

STATIC FUNCTION _chksendltr()

IF NEWUPS->SENDLTR='Y'
   return .T.
else
	return FALSE
endif
return FALSE



STATIC FUNCTION _BLBYBLREPS(dDateFrom,cSm)
	LOCAL GETLIST:={}, GETOPTIONS
	local xStatus, dDateto:=DATE(), nPage:=0
	local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
	LOCAL aLATTR:=ARRAY(GRA_AL_COUNT) , nXcoord:=0 , cCell, cWork , cHome
	TOP_MAR(2)
	BOT_MAR(2)
	PAGE_LEN(62)
	aLATTR[GRA_AL_COLOR]:=GRA_CLR_BLACK
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE

	NEWUPS->(DBCLEARFILTER())

	ORDSETFOCUS('NEWUPDT')
	NEWUPS->(DBGOTOP())


	/// DIALOG
	@ 10,0 DCSAY 'Enter Salesperson to print for or leave blank for all. ' get cSm picture '@!'
	@ 11,0 DCSAY 'Enter the begining date to print for:' get dDatefrom
	@ 12,0 DCSAY 'Enter the ending date to print for:  ' get dDateto
	DCREAD GUI MODAL TO xStatus ADDBUTTONS FIT OPTIONS GETOPTIONS TITLE 'Sales Process Reports ' EVAL{|o| SETAPPWINDOW(o)}
	IF !xStatus
		RETURN NIL
	ENDIF

	IF !Print_Choice('Sales Process Log', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
   ENDIF
   IF !PrintOn('Sales Process Log', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
   ENDIF

	IF nDefmode=8
		PAGE_LEN(32000)
	ENDIF

	IF !_UPPROXDATE(dDateFrom)
		DC_WINALERT('No Ups in this date range!')
		RETURN NIL
	ENDIF



   SELECT NEWUPS
	IF !EMPTY(cSm)
		_BLBYBLHDR(@nPage,dDateFrom,dDateto,nDefmode)
	   DO WHILE !NEWUPS->(EOF())
			IF NEWUPS->DATEIN <= dDateto
			 IF NEWUPS->SALESMAN=cSm
				_prnttherecord()
				IF DCPAGEEJECT()
			      _BLBYBLHDR(@nPage,dDateFrom,dDateto,nDefmode)
		      ENDIF
			 ENDIF
			ENDIF
			SKIP ALIAS NEWUPS
			IF NEWUPS->(EOF())
				EXIT
			ENDIF
		ENDDO
	ENDIF


	IF EMPTY(cSm)
   	_BLBYBLHDR(@nPage,dDateFrom,dDateto,nDefmode)
		DO WHILE !NEWUPS->(EOF())
			IF NEWUPS->DATEIN <= dDateto
			 _prnttherecord()
			 IF DCPAGEEJECT()
			 _BLBYBLHDR(@nPage,dDateFrom,dDateto,nDefmode)
		    ENDIF
			ENDIF
			SKIP ALIAS NEWUPS
			IF NEWUPS->(EOF())
				EXIT
			ENDIF
		ENDDO
	ENDIF

 PRINTOFF(oPrinter)
RETURN nil

static function _prnttherecord()
		 @ DC_PRINTERROW()+1,0 DCPRINT SAY  NEWUPS->DATEIN
       @ DC_PRINTERROW(),10 DCPRINT SAY  NEWUPS->SALESMAN
       @ DC_PRINTERROW(),15 DCPRINT SAY  NEWUPS->Upstatus
       @ DC_PRINTERROW(),20 DCPRINT SAY  NEWUPS->UPTYPE
       @ DC_PRINTERROW(),25 DCPRINT SAY  NEWUPS->ORIGSOURCE
       @ DC_PRINTERROW(),30 DCPRINT SAY  SUBSTR(NEWUPS->LASTNAME,1,20)
       @ DC_PRINTERROW(),55 DCPRINT SAY NEWUPS->APPOINT  PICTURE 'Y'
		 IF _CHK4SHOW(NEWUPS->UPID)
			 @ DC_PRINTERROW(),60 DCPRINT SAY 'Y'
			else
			 @ DC_PRINTERROW(),60 DCPRINT SAY 'N'
		 ENDIF
       @ DC_PRINTERROW(),65 DCPRINT SAY NEWUPS->DEMOED PICTURE 'Y'
       @ DC_PRINTERROW(),70 DCPRINT SAY NEWUPS->TOED   PICTURE 'Y'
       @ DC_PRINTERROW(),75 DCPRINT SAY NEWUPS->WRITEUP PICTURE 'Y'
       @ DC_PRINTERROW(),80 DCPRINT SAY NEWUPS->SOLD    PICTURE 'Y'
		 @ DC_PRINTERROW(),85 DCPRINT SAY NEWUPS->NEWUSED
		 @ DC_PRINTERROW(),90 DCPRINT SAY NEWUPS->ADSOURCE
		 @ DC_PRINTERROW(),100 DCPRINT SAY NEWUPS->INTEREST
		 @ DC_PRINTERROW(),110 DCPRINT SAY NEWUPS->TOMGR
		 @ DC_PRINTERROW(),115 DCPRINT SAY NEWUPS->BEBACK3
		 @ DC_PRINTERROW(),123 DCPRINT SAY NEWUPS->BEBACK1
	return nil

STATIC FUNCTION _BLBYBLHDR(nPage,dDateFrom,dDateto,nDefmode)
		 nPage++
		IF nDefmode <> 8
		 @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Sales Process Review For '+ dtoc(dDateFrom) +' to ' +dtoc(dDateTo)+'  Page '+ alltrim(str(nPage))
		endif
		@ DC_PRINTERROW()+1,0 DCPRINT SAY  'DateIn'
      @ DC_PRINTERROW(),10 DCPRINT SAY  'SP'
      @ DC_PRINTERROW(),15 DCPRINT SAY  'Stat'
      @ DC_PRINTERROW(),20 DCPRINT SAY  'Type'
      @ DC_PRINTERROW(),25 DCPRINT SAY  'SRC'
      @ DC_PRINTERROW(),30 DCPRINT SAY  'Lastname'
      @ DC_PRINTERROW(),55 DCPRINT SAY  'Appt'
      @ DC_PRINTERROW(),60 DCPRINT SAY  'Show'
      @ DC_PRINTERROW(),65 DCPRINT SAY 'Demo'
      @ DC_PRINTERROW(),70 DCPRINT SAY 'TOed'
      @ DC_PRINTERROW(),75 DCPRINT SAY 'W/Up'
      @ DC_PRINTERROW(),80 DCPRINT SAY 'Sold'
		@ DC_PRINTERROW(),85 DCPRINT SAY 'N/U'
		@ DC_PRINTERROW(),90 DCPRINT SAY 'AdSrc'
		@ DC_PRINTERROW(),100 DCPRINT SAY 'Interest'
		@ DC_PRINTERROW(),110 DCPRINT SAY 'TOMgr'
		@ DC_PRINTERROW(),115 DCPRINT SAY 'OrigDtIn'
		@ DC_PRINTERROW(),123 DCPRINT SAY 'Beback1'
		skipaline()
	return nil


static function _UPPROXDATE(dDateFrom)
	SELECT NEWUPS
	ORDSETFOCUS('NEWUPDT')
	SET SOFTSEEK ON
	GOTO TOP
	***"HROMASDT")
	SEEK dDateFrom
	IF found()
		RETURN TRUE
   	ELSE
		SEEK dDatefrom+1
		IF found()
		 RETURN TRUE
	    ELSE
		 SEEK dDatefrom+2
		 IF found()
		  RETURN TRUE
	     ELSE
		  SEEK dDatefrom+3
		  IF found()
			return TRUE
		  ENDIF
		 ENDIF
	   ENDIF
	 ENDIF
RETURN .F.

/*
FUNCTION MAINUP(lReshow)
	LOCAL GETLIST:={},oRecord
	LOCAL GETOPTIONS , xStatus , Aotoolbar , lFound , cMiddleinit , clAfterDel:=' ', cBillStat:=''
	LOCAL MCON,  aUpChars:={},aUpNums:={},aUpdates:={} , aPres, oBrowse , aAppts:={} , aUpBoolean:={}
	local cUpstat,cCodeath,cVehicleID,nMSRP,nTissue , dDatesold , cNewUsed , cFranchise:='', nDowndeal:=0 ,nCurrUpType:=0
	Local nRec:=NEWUPS->(recno()) ,oConfig1,oConfig2 ,oUpinfo,oVehinfo,oUpstat,oProgInfo
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE SAYFONT '12.Courier New' getfont '12.Courier New'  NOTABSTOP
	MCON:='Y'

	oRecord:=NEWUPS->(DC_DBRECORD(): NEW())
	NEWUPS->(DC_DBSCATTER(oRecord))

	oConfig1 := DC_XbpPushButtonXPConfig():new()
	oConfig1:bitmapOffset := 5
	oConfig1:fgColorMouse := COLOR_BLACK
	oConfig1:bgColorMouse := COLOR_SILVER
	oConfig1:fgColor := COLOR_BLACK
	oConfig1:bgColor := COLOR_GREEN
	oConfig1:gradientStep := 10
	oConfig1:gradientReverse := .t.
	oConfig1:radius := 10

	oConfig2 := DC_XbpPushButtonXPConfig():new()
	oConfig2:bitmapOffset := 5
	oConfig2:fgColorMouse := COLOR_BLACK
	oConfig2:bgColorMouse := COLOR_YELLOW
	oConfig2:fgColor := COLOR_DARKBLUE
	oConfig2:bgColor := COLOR_WHITE
	oConfig2:gradientStep := 10
	oConfig2:gradientReverse := .t.
	oConfig2:radius := 10



	_initUPsCharArray(@aUpChars)
   _initUPsNumArray(@aUpNums)
	_initUPSDATEArray(@aUpDates)
	_initarray(@aUpBoolean,8,.f.)

	aUpChars[UPSC_CFRANCHISE]:=NEWUPS->FRANCHISE
	aUpDates[UPSD_DDATEIN] :=NEWUPS->DATEIN
	aUpDates[UPSD_DORIGDTIN] :=NEWUPS->BEBACK3
	aUpChars[UPSC_CAREACODE]:=NEWUPS->AREA
	aUpChars[UPSC_CUPID]:=NEWUPS->UPID
	aUpChars[UPSC_CFIRSTNAME]:=NEWUPS->FIRSTNAME
	aUpChars[UPSC_CLASTNAME]:=NEWUPS->LASTNAME
	aUpChars[UPSC_CSTREET]:=NEWUPS->STREET
	aUpChars[UPSC_CCITYSTATE]:=NEWUPS->CITYSTATE
	aUpChars[UPSC_STATE]:=NEWUPS->STATE
	aUpChars[UPSC_COUNTY]:=NEWUPS->COUNTY
	aUpChars[UPSC_CZIPCODE]:=NEWUPS->ZIPCODE
	aUpChars[UPSC_CVEHINTEREST]:=NEWUPS->INTEREST
	aUpNums[UPSN_NUPTYPE]:=NEWUPS->UPTYPE
	aUpChars[UPSC_CNEWUSED] :=NEWUPS->NEWUSED
	aUpChars[UPSC_CSALESPERSON]:=NEWUPS->SALESMAN
	aUpChars[UPSC_CCOMMENT]:=NEWUPS->COMMENTS
	aUpChars[UPSC_CCOMMENT2]:=NEWUPS->COMMENT2
	aUpNums[UPSN_NSOLDCODE]:=NEWUPS->SALECODE
	aUpChars[UPSC_CTOMGR]:=NEWUPS->TOMGR
	aUpChars[UPSC_CSTOCKNO]:=NEWUPS->STOCKNUM
	aUpChars[UPSC_CADSOURCE]:=NEWUPS->ADSOURCE
	aUpChars[UPSC_CLDEMOYN]:=NEWUPS->DEMO
	aUpChars[UPSC_CWORKPHONE]:=NEWUPS->WORKPHN
	aUpChars[UPSC_CCELLPHONE]:=NEWUPS->CELLPHONE
	aUpDates[UPSD_DLASTDTIN]:=NEWUPS->DATEIN
	aUpChars[UPSC_CDEMODSTKNO]:=NEWUPS->DEMOSTK
	aUpChars[UPSC_CEMAIL]:=NEWUPS->EMAIL
	aUpDates[UPSD_DTICKLEDATE]:=NEWUPS->TICKLE
	aUpDates[UPSD_DBEBACK1]:=NEWUPS->BEBACK1
	aUpDates[UPSD_DBEBACK2]:=NEWUPS->BEBACK2
	aUpChars[UPSC_CLBEBACKYN]:='N'
	aUpChars[UPSC_ORIGSOURCE]:=NEWUPS->ORIGSOURCE
	aUpChars[UPSC_PREFMETH]:=NEWUPS->PREFMETHOD

	IF EMPTY(aUpChars[UPSC_CUPID])
		DC_WINALERT('You must have an UP ID to maintain. Please put one in on the browser.')
      RETURN NIL
	ENDIF

	cBillstat:=NEWUPS->BILLSTAT
	IF cBillstat='D'
		aUpBoolean[UPSL_DELIVERED]:=.T.
	ENDIF
	aUpBoolean[UPSL_APPOINT]=NEWUPS->APPOINT
	aUpBoolean[UPSL_DEMOED]:=NEWUPS->DEMOED
	aUpBoolean[UPSL_WRITEUP]:=NEWUPS->WRITEUP
	aUpBoolean[UPSL_TOED]:=NEWUPS->TOED
	aUpBoolean[UPSL_SOLD]:=NEWUPS->SOLD

	cVehicleID:=SPACE(17)
	nMSRP:=nTissue:=0
	nCurrUpType:=NEWUPS->UPTYPE       /// set up to look for a change

	IF NEWUPS->NEWUSED='N'
		cNewUsed:='N'
	ELSE
		cNewUsed :='U'
	ENDIF
	cUpstat:=NEWUPS->upstatus
	nDowndeal:=NEWUPS->downdeal
	cCodeath:=NEWUPS->causedead
	dDatesold:=NEWUPS->saledate
	cMiddleinit:=NEWUPS->middleinit
	clAfterdel:=NEWUPS->SENDLTR
	IF clAfterdel='Y'
		aUpBoolean[UPSL_FOLLOWED]:=.T.
	ENDIF
	@ 0,1 DCGROUP oUpStat Caption 'Prospect Status' Size 140,7.2

	@ 1,1 DCSAY 'Up Id' GET aUpChars[UPSC_CUPID]  SAYCOLOR 1,BD_LEMONCREME ; //editprotect{|| TRUE } ;
		PARENT oUpStat

	@ 2,1 DCSAY 'Status  (A)ctive (R)escue (F)uture (S)old (D)ead' get cUpstat valid{|| editupstat(@cUpstat,@aUpBoolean,@GETLIST)};
			SAYCOLOR 1,BD_LEMONCREME PICTURE '!'    PARENT oUpStat  GETID 'UPSTAT'

	@ 2,90 DCSAY 'Stock Number If Sold' Get aUpChars[UPSC_CSTOCKNO]  picture '@!' WHEN{ || IIF(cUpstat='S',.t. ,.f.)};
		valid{|| DC_GETREFRESH(GETLIST),.T.} SAYCOLOR 9,0 PARENT oUpStat
	@ 3,1 DCSAY 'Dead deal Code (C)redit (B)ought Elsewhere (I)nventory (N)o Response (O)ther (U)psideDown' get cCodeath ;
		when {|| iif(cUpstat='D',.t.,.f.)};
		valid{|| iif(cCodeath $('CBIONU'),.t.,.f.)} ;
		picture '!' SAYCOLOR 9,0   PARENT oUpStat

	@ 4,1 DCSAY '                             OrigDateIn' get aUpDates[UPSD_DORIGDTIN] EDITPROTECT{|| TRUE }  PARENT oUpStat


	@ 5,1 DCSAY '              Enter Next Follow up Date' GET NEWUPS->TICKLE  WHEN{|| IIF( cUpstat $('AFR'),.t. ,.f. )}   SAYCOLOR 9,0    PARENT oUpStat


	@ 6,1 DCSAY 'Billing Status (I)nProgress (D)elivered' get cBillstat PICTURE '@!' valid {|| _chkbillstat(cBillstat,@aUpBoolean,@getlist) };
		when {|| iif(cUpstat=='S',.t.,.f.)};
		picture '!' SAYCOLOR 9,0 PARENT oUpStat


	@ 7,1 DCGROUP oUpinfo Caption 'Prospect Demographic' Size 140,10

	@ 1,1 DCSAY '      Firstname' Get aUpChars[UPSC_CFIRSTNAME]  Saycolor 9,0 picture '@!' PARENT oUpInfo
	@ 1,70 DCSAY 'Middle Initial' get cMiddleInit Saycolor 9,0 picture '@!'                PARENT oUpInfo
	@ 2,1 DCSAY '       Lastname' Get aUpChars[UPSC_CLASTNAME]  Saycolor 9,0  picture '@!' PARENT oUpInfo
	@ 3,1 DCSAY '   Enter Street' Get aUpChars[UPSC_CSTREET]  Saycolor 9,0  picture '@!'   PARENT oUpInfo
	@ 3,70 DCSAY 'Enter Zipcode' Get  aUpChars[UPSC_CZIPCODE] SAYCOLOR 9,0                PARENT oUpInfo ;
		VALID{|| SEEKZIP(@aUpChars),DC_GETREFRESH(GETLIST),.T.};
		POPUP{|| BROWGETZIP(@aUpChars),DC_GETREFRESH(GETLIST)};
		POPCAPTION 'ZipSearch ' POPWIDTH 100 POPFONT '10.Courier New' POPSTYLE 1

	@ 4,1 DCSAY '      Citystate' Get aUpChars[UPSC_CCITYSTATE]  Saycolor 9,0   picture '@!' PARENT oUpInfo
	@ 4,65 DCSAY 'State' get aUpChars[UPSC_STATE]                                            PARENT oUpInfo
	@ 4,80 DCSAY 'County' get aUpChars[UPSC_COUNTY]                                            PARENT oUpInfo

	@ 5,1 DCSAY '   Last Date In' Get aUpDates[UPSD_DDATEIN]                                 PARENT oUpInfo

	@ 5,45 DCSAY 'Beback1' Get aUpDates[UPSD_DBEBACK1]  WHEN{|| aUpNums[UPSN_NUPTYPE] < 3}   PARENT oUpInfo
	@ 5,70 DCSAY 'Beback2' Get aUpDates[UPSD_DBEBACK2]  WHEN{|| IIF(!EMPTY(aUpDates[UPSD_DBEBACK1]),.T. ,.F.)} PARENT oUpInfo

	@ 6,1 DCSAY '          Email' Get aUpChars[UPSC_CEMAIL] Saycolor 9,0   picture '@!' valid{|| chkemailaddress(@aUpChars[UPSC_CEMAIL],.f.)} PARENT oUpInfo
	@ 7,1 DCSAY '           Area' get aUpChars[UPSC_CAREACODE]        PARENT oUpInfo   PICTURE '999' SAYCOLOR 9,0
	@ 7,30 DCSAY 'Main Phone#' get aUpChars[UPSC_CUPID]  SAYCOLOR 9,0   PARENT oUpInfo PICTURE '@R 999-9999'
	@ 7,60 DCSAY 'AltPhone#'  Get aUpChars[UPSC_CWORKPHONE]  Saycolor 9,0 PARENT oUpInfo   PICTURE '@R 999-999-9999-AAAA'
	@ 7,100 DCSAY 'Cell #' get aUpChars[UPSC_CCELLPHONE]  Saycolor 9,0 PARENT oUpInfo       PICTURE '@R 999-999-9999'
	@ 8,1 DCSAY '     Contact Method [A]ll Ok [H]omephone [C]ell [W]ork [E]mail [D]ONOTCALL!' get aUpChars[UPSC_PREFMETH] PARENT oUpInfo;
		PICTURE '@!' ;
		VALID {|| IIF( aUpChars[UPSC_PREFMETH] $('AHCWED'),.T. ,.F. )} ;
		saycolor GRA_CLR_DARKBLUE,GRA_CLR_WHITE


	@ 17,1 DCGROUP oVehinfo CAPTION 'Vehicle Interest' size 140,6.5
	@ 1,1 DCSAY '                                                    Vehicle Interest' GET aUpChars[UPSC_CVEHINTEREST] picture '@!'  PARENT oVehInfo;
		SAYCOLOR 9,0;
		VALID{||CHKINTR(@aUpChars[UPSC_CVEHINTEREST],@aUpChars[UPSC_CFRANCHISE]),dc_getrefresh(getlist),.T.}

	@ 2,1 DCSAY '                Source (F)loor Traffic  (I)nternet Lead  (P)hone Up ' get aUpChars[UPSC_ORIGSOURCE] VALID{|| IIF( aUpChars[UPSC_ORIGSOURCE] $('FIP'),.T. ,.F. )}  ;
		PICTURE '@!' SAYCOLOR 9,0  ;
		GETTOOLTIP('You must enter the source of this Up. This will be permanently stored as the original source.')  PARENT oVehInfo

	@ 3,1 DCSAY 'UP Type 1=NewCust 2=Cust/Refer 3=Beback/1 4=Beback/2. 5=iNet 6=Phone' GET aUpNums[UPSN_NUPTYPE] PICTURE '9'   PARENT oVehInfo;
		 VALID{|| _CHK4TYPECHANGE(nCurrUptype,aUpNums,@aUpDates),DC_GETREFRESH(GETLIST),.t.} ;
		 RANGE 1,6  SAYCOLOR 9,0  ;
		 GETTOOLTIP('For floor traffic always use 123 or 4. For iNet or phone sources always use 5 or 6 UNTIL the up ; actually comes in. THEN change it to a 1 or 2 and the up will be eligible in the up count for the period.')
	@ 4,1 DCSAY '                      Enter N For New Vehicle Or U For Used vehicle.' GET aUpChars[UPSC_CNEWUSED] PICTURE '@!' VALID{|| CHKNU(aUpChars[UPSC_CNEWUSED])} SAYCOLOR 9,0   ;
		PICTURE '@!' PARENT oVehInfo
	@ 5,1 DCSAY '                      Advertising Source or Press Enter for Choices.' GET aUpChars[UPSC_CADSOURCE]  PARENT oVehInfo;
		SAYCOLOR 9,0;
		PICTURE '@!' ;
		valid {|| aslk(@aUpChars[UPSC_CADSOURCE]),dc_getrefresh(getlist),.T.}

	@ 23.5,1 DCGROUP oProgInfo CAPTION 'Prospect Progress' size 140,8.3

	@ 1,1 DCSAY '      Comments'  Get aUpChars[UPSC_CCOMMENT]  Saycolor 9,0 PARENT oProgInfo
	@ 2,1 DCSAY ' More Comments' Get aUpChars[UPSC_CCOMMENT2]  Saycolor 9,0 PARENT oProgInfo
	@ 3,1 DCSAY '   T.O.Manager'  Get  aUpChars[UPSC_CTOMGR] picture '@!' Saycolor 9,0 PARENT oProgInfo;
		VALID{|| IIF( !EMPTY(aUpChars[UPSC_CTOMGR]) ,;
		aUpBoolean[UPSL_TOED]:=.T. ,aUpBoolean[UPSL_TOED]:=.F. ),DC_GETREFRESH(GETLIST),.T.}

	@ 4,1 DCSAY ' Demo Ride Y/N' Get aUpChars[UPSC_CLDEMOYN] PARENT oProgInfo Picture '@! L'  Saycolor 9,0 VALID{|| IIF(aUpChars[UPSC_CLDEMOYN]='Y' ,;
		aUpBoolean[UPSL_DEMOED]:=.T. ,aUpBoolean[UPSL_DEMOED]:=.F. ),DC_GETREFRESH(GETLIST),.T.}
	@ 4,40 DCSAY 'Stk# Demonstrated ' Get aUpChars[UPSC_CDEMODSTKNO]  picture '@!' WHEN{|| IIF(aUpChars[UPSC_CLDEMOYN]='Y' ,.T. ,.F. )};
		SAYCOLOR 9,0 PARENT oProgInfo

	@ 5,1 DCSAY 'If Down Deal Enter Down Code 0,1,2,3,4,5,6' Get nDowndeal PARENT oProgInfo Saycolor 9,0 Picture '9' Range 0,6

	@ 6,1 DCSAY 'Appointment' get aUpBoolean[UPSL_APPOINT]  EDITPROTECT{|| TRUE }  NOTABSTOP PICTURE 'Y' PARENT oProgInfo ;
		  SAYCOLOR{|| IIF(aUpBoolean[UPSL_APPOINT] ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

	@ 6,24 DCSAY 'Showed Up' get aUpBoolean[UPSL_SHOWED] EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo ;
		  SAYCOLOR {|| _CHK4SHOWCLR(aUpChars[UPSC_CUPID])}

	@ 6,44 DCSAY 'Demo Ride' get aUpBoolean[UPSL_DEMOED]    EDITPROTECT{|| TRUE }  NOTABSTOP   PICTURE 'Y' PARENT oProgInfo ;
		SAYCOLOR{|| IIF(aUpBoolean[UPSL_DEMOED] ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

	@ 6,62 DCSAY 'TOed to Manager' get aUpBoolean[UPSL_TOED]  EDITPROTECT{|| TRUE } NOTABSTOP  PICTURE 'Y' PARENT oProgInfo;
		SAYCOLOR{|| IIF(aUpBoolean[UPSL_TOED] ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

	@ 6,90 DCSAY 'Written Up'  get aUpBoolean[UPSL_WRITEUP]    EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo;
		SAYCOLOR{|| IIF(aUpBoolean[UPSL_WRITEUP] ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

	@ 6,110 DCSAY 'Sold'  get aUpBoolean[UPSL_SOLD]            EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo;
		SAYCOLOR{|| IIF(aUpBoolean[UPSL_SOLD] ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

   @ 6,123 DCSAY 'Delivered'  get aUpBoolean[UPSL_DELIVERED]       EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo ;
		SAYCOLOR{|| IIF(aUpBoolean[UPSL_DELIVERED] ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,GRA_CLR_PALEGRAY} )}

	@ 7.1 DCSAY 'After Delivery Call Made Y/N' get clAfterdel PICTURE '@! L' WHEN {|| IIF(cBillstat='D',.T.,.F. )} PARENT oProgInfo   ;
		VALID{|| IIF(clAfterdel='Y' ,aUpBoolean[UPSL_FOLLOWED]:=.T. ,aUpBoolean[UPSL_FOLLOWED]:=.F. ),DC_GETREFRESH(GETLIST)}

	@ 7,50 DCSAY 'After Delivery Call' get aUpBoolean[UPSL_FOLLOWED] EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo ;
		SAYCOLOR{|| IIF(aUpBoolean[UPSL_FOLLOWED] ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,GRA_CLR_PALEGRAY} )}

	@ 32,0 DCTOOLBAR AoToolBar                               ;
		SIZE 140,2


@ 2,120 DCPUSHBUTTONXP CAPTION 'Record Be Back' ;
	 SIZE 20,2 ;
	 ACTION {|| _MAKEBEBACK(@aUpDates,aUpNums),DC_GETREFRESH(GETLIST)} ;
	 CONFIG oConfig1  ;
	 BITMAP 7755      ;
	 PARENT oUpinfo

@ 4,120 DCPUSHBUTTONXP CAPTION 'Send Record to Vin ' ;
	 SIZE 20,2 ;
	 ACTION {|| UpdateUpRec(aUpChars,aUpNums,aUpdates,aUpBoolean,nDowndeal,dDatesold,cMiddleinit,cCoDeath,cUpstat,clAfterdel,cBillStat),_senduptovin(NEWUPS->(RECNO()))};
	 CONFIG oConfig2  ;
	 BITMAP 8062 ;
	 PARENT oUpinfo


IF MANAGER()
	DCADDBUTTON CAPTION 'Payment;Calculator'                              ;
		SIZE 12                                              ;
		ACTION {|| QUICK(0,0,0,0,0,0,0,0,0,0,.T.,0)  };
		PARENT AoToolBar      ;
		HIDE {|| IF(!MANAGER(),.T.,.F.)}
ENDIF

DCADDBUTTON CAPTION 'Comments '                              ;
	SIZE 12                                              ;
	ACTION {|| UPNOTES()  };
	PARENT AoToolBar


DCADDBUTTON CAPTION 'Print UPSheet '                              ;
	SIZE 12                                              ;
	ACTION {|| UpdateUpRec(aUpChars,aUpNums,aUpdates,aUpBoolean,nDowndeal,dDatesold,cMiddleinit,cCoDeath,cUpstat,clAfterdel,cBillStat),;
	PRINTUPSHEET(NEWUPS->(RECNO()))};
	PARENT AoToolBar

DCADDBUTTON CAPTION 'Desking;Tool '                              ;
		SIZE 12                                              ;
		ACTION {||upquote(@aUpBoolean),DC_GETREFRESH(GETLIST)};
		PARENT AoToolBar


DCADDBUTTON CAPTION 'Buyers Order '                              ;
		SIZE 12                                              ;
		ACTION {|| UpdateUpRec(aUpChars,aUpNums,aUpdates,aUpBoolean,nDowndeal,dDatesold,cMiddleinit,cCoDeath,cUpstat,clAfterdel,cBillStat),;
		  BuyersOrder(NEWUPS->(RECNO()),.f.,@aUpBoolean),DC_GETREFRESH(GETLIST)};
		PARENT AoToolBar

IF MANAGER()
  DCADDBUTTON CAPTION 'Final Buyers; Order '                              ;
		SIZE 15                                              ;
		ACTION {|| UpdateUpRec(aUpChars,aUpNums,aUpdates,aUpBoolean,nDowndeal,dDatesold,cMiddleinit,cCoDeath,cUpstat,clAfterdel,cBillStat),;
		BuyersOrder(NEWUPS->(RECNO()),.t.,@aUpBoolean),DC_GETREFRESH(GETLIST)};
		PARENT AoToolBar
ENDIF

DCADDBUTTON CAPTION 'Appraisal Form '                              ;         /// this will update Newups appraisal record
		SIZE 12                                              ;
		ACTION {|| UpdateUpRec(aUpChars,aUpNums,aUpdates,aUpBoolean,nDowndeal,dDatesold,cMiddleinit,cCoDeath,cUpstat,clAfterdel,cBillStat),;
		_appraiseform(NEWUPS->(RECNO()),aUpChars[UPSC_CUPID],aUpChars[UPSC_CSALESPERSON],oRecord),;
		NEWUPS->(DBGOTO(oRecord:Record_Number)),;
	   NEWUPS->(DC_DBGATHER(oRecord))};
		PARENT AoToolBar

DCADDBUTTON CAPTION 'Review eMails '                              ;
		SIZE 15                                              ;
		ACTION {|| BROWEMAILLOG('',NEWUPS->EMAIL)} ;
		PARENT AoToolBar


DCADDBUTTON CAPTION 'Exit '                              ;
	SIZE 10                                              ;
	ACTION {||DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)};
	PARENT AoToolBar

DCREAD GUI TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA ;
		EVAL {|| SETAPPFOCUS(DC_GETOBJECT(GETLIST,'UPSTAT'))};
	 	FIT ADDBUTTONS OPTIONS GETOPTIONS;
	  TITLE 'Up Record Maintainance'
	  //	EVAL{|o| setappwindow(o)}

IF !XSTATUS
 RETURN NIL
ENDIF

UpdateUpRec(aUpChars,aUpNums,aUpdates,aUpBoolean,nDowndeal,dDatesold,cMiddleinit,cCoDeath,cUpstat,clAfterdel,cBillStat)
//DBCLEARSCOPE()
RETURN NIL     */




/// new version recordobject
FUNCTION XMAINUP()
	LOCAL GETLIST:={},oRecord
	LOCAL GETOPTIONS , xStatus , Aotoolbar , nCurrUptype:=0
	LOCAL aPres, oBrowse , aAppts:={} , lDelivered:=.f.
	Local oConfig1,oConfig2 ,oUpinfo,oVehinfo,oUpstat,oProgInfo,oProgInfo2
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE SAYFONT '12.Courier New' getfont '12.Courier New'  NOTABSTOP

	oRecord:=NEWUPS->(DC_DBRECORD(): NEW())
	NEWUPS->(DC_DBSCATTER(oRecord))

	oConfig1 := DC_XbpPushButtonXPConfig():new()
	oConfig1:bitmapOffset := 5
	oConfig1:fgColorMouse := COLOR_BLACK
	oConfig1:bgColorMouse := COLOR_SILVER
	oConfig1:fgColor := COLOR_BLACK
	oConfig1:bgColor := COLOR_GREEN
	oConfig1:gradientStep := 10
	oConfig1:gradientReverse := .t.
	oConfig1:radius := 10

	oConfig2 := DC_XbpPushButtonXPConfig():new()
	oConfig2:bitmapOffset := 5
	oConfig2:fgColorMouse := COLOR_BLACK
	oConfig2:bgColorMouse := COLOR_YELLOW
	oConfig2:fgColor := COLOR_DARKBLUE
	oConfig2:bgColor := COLOR_WHITE
	oConfig2:gradientStep := 10
	oConfig2:gradientReverse := .t.
	oConfig2:radius := 10



	IF EMPTY(oRecord:UPID)
		DC_WINALERT('You must have an UP ID to maintain. Please put one in on the browser.')
      RETURN NIL
	ENDIF

	IF oRecord:Billstat='D'
	 	lDelivered:=.T.
	ENDIF

	nCurrUpType:=NEWUPS->UPTYPE       /// set up to look for a change

	//cMiddleinit:=NEWUPS->middleinit
	//clAfterdel:=NEWUPS->SENDLTR
	@ 0,1 DCGROUP oUpStat Caption 'Prospect Status' Size 140,7.2

	@ 1,1 DCSAY 'Up Id' GET oRecord:UPID  SAYCOLOR 1,BD_LEMONCREME ; //editprotect{|| TRUE } ;
		PARENT oUpStat

	@ 2,1 DCSAY 'Status  (A)ctive (R)escue (F)uture (S)old (D)ead' get oRecord:Upstatus valid{|| editupstat(@oRecord)};
			SAYCOLOR 1,BD_LEMONCREME PICTURE '!'    PARENT oUpStat  GETID 'UPSTAT'

	@ 2,90 DCSAY 'Stock Number If Sold' Get oRecord:STOCKNUM  picture '@!' WHEN{ || IIF(oRecord:Upstatus='S',.t. ,.f.)};
		valid{|| DC_GETREFRESH(GETLIST),.T.} SAYCOLOR 9,0 PARENT oUpStat
	@ 3,1 DCSAY 'Dead deal Code (C)redit (B)ought Elsewhere (I)nventory (N)o Response (O)ther (U)psideDown' get oRecord:causedead ;
		when {|| iif(oRecord:Upstatus='D',.t.,.f.)};
		valid{|| iif(oRecord:Causedead $('CBIONU'),.t.,.f.)} ;
		picture '!' SAYCOLOR 9,0   PARENT oUpStat

	@ 4,1 DCSAY '                             OrigDateIn' get oRecord:beback3 EDITPROTECT{|| TRUE }  PARENT oUpStat


	@ 5,1 DCSAY '              Enter Next Follow up Date' GET oRecord:TICKLE  WHEN{|| IIF( oRecord:Upstatus $('AFR'),.t. ,.f. )}   SAYCOLOR 9,0    PARENT oUpStat


	@ 6,1 DCSAY 'Billing Status (I)nProgress (D)elivered' get oRecord:Billstat PICTURE '@!' valid {|| _chkbillstat(oRecord,@getlist) };
		when {|| iif(oRecord:Upstatus='S',.t.,.f.)};
		picture '!' SAYCOLOR 9,0 PARENT oUpStat


	@ 7,1 DCGROUP oUpinfo Caption 'Prospect Demographic' Size 140,10

	@ 1,1 DCSAY '      Firstname' Get oRecord:FIRSTNAME  Saycolor 9,0 picture '@!' PARENT oUpInfo
	@ 1,70 DCSAY 'Middle Initial' get oRecord:MiddleInit Saycolor 9,0 picture '@!'                PARENT oUpInfo
	@ 2,1 DCSAY '       Lastname' Get oRecord:LASTNAME  Saycolor 9,0  picture '@!' PARENT oUpInfo
	@ 3,1 DCSAY '         Street' Get oRecord:STREET  Saycolor 9,0  picture '@!'   PARENT oUpInfo
	@ 3,70 DCSAY '       Zipcode' Get  oRecord:ZIPCODE SAYCOLOR 9,0                PARENT oUpInfo ;
		VALID{|| SEEKZIP(@oRecord),DC_GETREFRESH(GETLIST),.T.};
		POPUP{|| BROWGETZIP(@oRecord),DC_GETREFRESH(GETLIST)};
		POPCAPTION 'ZipSearch ' POPWIDTH 100 POPFONT '10.Courier New' POPSTYLE 1

	@ 4,1 DCSAY '      Citystate' Get oRecord:CITYSTATE  Saycolor 9,0   picture '@!' PARENT oUpInfo
	@ 4,61 DCSAY 'State' get oRecord:STATE                                            PARENT oUpInfo
	@ 4,81 DCSAY ' County' get oRecord:COUNTY                                            PARENT oUpInfo

	@ 5,1 DCSAY '   Last Date In' Get oRecord:DATEIN                                 PARENT oUpInfo

	@ 5,50 DCSAY 'Beback1' Get oRecord:BEBACK1  WHEN{|| oRecord:UPTYPE < 3}   PARENT oUpInfo
	@ 5,81 DCSAY 'Beback2' Get oRecord:BEBACK2  WHEN{|| IIF(!EMPTY(oRecord:BEBACK1),.T. ,.F.)} PARENT oUpInfo

	@ 6,1 DCSAY '          Email' Get oRecord:EMAIL Saycolor 9,0   picture '@!' valid{|| chkemailaddress(@oRecord:EMAIL,.f.)} PARENT oUpInfo
	@ 7,1 DCSAY '           Area' get oRecord:AREA        PARENT oUpInfo   PICTURE '999' SAYCOLOR 9,0
	@ 7,30 DCSAY 'Main Phone#' get oRecord:UPID  SAYCOLOR 9,0   PARENT oUpInfo PICTURE '@R 999-9999'
	@ 7,60 DCSAY 'AltPhone#'  Get oRecord:WORKPHN  Saycolor 9,0 PARENT oUpInfo   PICTURE '@R 999-999-9999-AAAA'
	@ 7,100 DCSAY 'Cell #' get oRecord:CELLPHONE  Saycolor 9,0 PARENT oUpInfo       PICTURE '@R 999-999-9999'
	@ 8,1 DCSAY '     Contact Method [A]ll Ok [H]omephone [C]ell [W]ork [E]mail [D]ONOTCALL!' get oRecord:PREFMETHOD PARENT oUpInfo;
		PICTURE '@!' ;
		VALID {|| IIF( oRecord:PREFMETHOD $('AHCWED'),.T. ,.F. )} ;
		saycolor GRA_CLR_DARKBLUE,GRA_CLR_WHITE


	@ 17,1 DCGROUP oVehinfo CAPTION 'Vehicle Interest' size 140,6.5
	@ 1,1 DCSAY '                                                    Vehicle Interest' GET oRecord:INTEREST picture '@!'  PARENT oVehInfo;
		SAYCOLOR 9,0;
		VALID{||CHKINTR(@oRecord:INTEREST,@oRecord:FRANCHISE),dc_getrefresh(getlist),.T.}

	@ 2,1 DCSAY '                Source (F)loor Traffic  (I)nternet Lead  (P)hone Up ' get oRecord:ORIGSOURCE VALID{|| IIF(oRecord:ORIGSOURCE $('FIP'),.T. ,.F. )}  ;
		PICTURE '@!' SAYCOLOR 9,0  ;
		GETTOOLTIP('You must enter the source of this Up. This will be permanently stored as the original source.')  PARENT oVehInfo

	@ 3,1 DCSAY 'UP Type 1=NewCust 2=Cust/Refer 3=Beback/1 4=Beback/2. 5=iNet 6=Phone' GET oRecord:UPTYPE PICTURE '9'   PARENT oVehInfo;
		 VALID{|| _CHK4TYPECHANGE(nCurrUptype,@oRecord),DC_GETREFRESH(GETLIST),.t.} ;
		 RANGE 1,6  SAYCOLOR 9,0  ;
		 GETTOOLTIP('For floor traffic always use 123 or 4. For iNet or phone sources always use 5 or 6 UNTIL the up ; actually comes in. THEN change it to a 1 or 2 and the up will be eligible in the up count for the period.')
	@ 4,1 DCSAY '                      Enter N For New Vehicle Or U For Used vehicle.' GET oRecord:NEWUSED PICTURE '@!' VALID{|| CHKNU(oRecord:NEWUSED)} SAYCOLOR 9,0   ;
		PICTURE '@!' PARENT oVehInfo
	@ 5,1 DCSAY '                      Advertising Source or Press Enter for Choices.' GET oRecord:ADSOURCE  PARENT oVehInfo;
		SAYCOLOR 9,0;
		PICTURE '@!' ;
		valid {|| aslk(@oRecord:ADSOURCE),dc_getrefresh(getlist),.T.}

	@ 23.5,1 DCGROUP oProgInfo CAPTION 'Other Prospect Info' size 70,8.3

	@ 1,1 DCSAY '      Comments'  Get oRecord:COMMENTS  Saycolor 9,0 PARENT oProgInfo
	@ 2,1 DCSAY ' More Comments' Get oRecord:COMMENT2  Saycolor 9,0 PARENT oProgInfo
	@ 3,1 DCSAY '   T.O.Manager'  Get  oRecord:TOMGR picture '@!' Saycolor 9,0 PARENT oProgInfo;
		VALID{|| IIF( !EMPTY(oRecord:TOMGR) ,;
		oRecord:TOED:=.T. ,oRecord:TOED:=.F. ),DC_GETREFRESH(GETLIST),.T.}

	@ 4,1 DCSAY ' Demo Ride Y/N' Get oRecord:DEMO PARENT oProgInfo Picture '@! L'  Saycolor 9,0 VALID{|| IIF(oRecord:DEMO='Y' ,;
		oRecord:DEMOED:=.T. ,oRecord:DEMOED:=.F. ),DC_GETREFRESH(GETLIST),.T.}
	@ 4,30 DCSAY 'Stk# Demonstrated ' Get oRecord:DEMOSTK  picture '@!' WHEN{|| IIF(oRecord:DEMO='Y' ,.T. ,.F. )};
		SAYCOLOR 9,0 PARENT oProgInfo

	@ 5,1 DCSAY 'If Down Deal Enter Down Code 0,1,2,3,4,5,6' Get oRecord:Downdeal PARENT oProgInfo Saycolor 9,0 Picture '9' Range 0,6
	@ 6,1 DCSAY 'After Delivery Call Made Y/N              ' get oRecord:SENDLTR PICTURE '@! L' Saycolor 9,0;
		 WHEN {|| IIF(oRecord:Billstat='D',.T.,.F. )} PARENT oProgInfo   ;


	@ 23.5,71 DCGROUP oProgInfo2 CAPTION 'Process Progress' size 70,8.3


	@ 1,1 DCSAY '         Appointment' get oRecord:APPOINT  EDITPROTECT{|| TRUE }  NOTABSTOP PICTURE 'Y' PARENT oProgInfo2 ;
		  SAYCOLOR{|| IIF(oRecord:APPOINT ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

	/*
	@ 6,24 DCSAY 'Showed Up' get lShowed EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo ;
		  SAYCOLOR {|| _CHK4SHOWCLR(oRecord:UPID)}    */

	@ 2,1 DCSAY '           Demo Ride' get oRecord:DEMOED    EDITPROTECT{|| TRUE }  NOTABSTOP   PICTURE 'Y' PARENT oProgInfo2 ;
		SAYCOLOR{|| IIF(oRecord:DEMOED ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

	@ 3,1 DCSAY '     TOed to Manager' get oRecord:TOED  EDITPROTECT{|| TRUE } NOTABSTOP  PICTURE 'Y' PARENT oProgInfo2;
		SAYCOLOR{|| IIF(oRecord:TOED ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

	@ 4,1 DCSAY '          Written Up'  get oRecord:WRITEUP    EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo2;
		SAYCOLOR{|| IIF(oRecord:WRITEUP ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

	@ 5,1 DCSAY '                Sold'  get oRecord:SOLD            EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo2;
		SAYCOLOR{|| IIF(oRecord:SOLD ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

   @ 6,1 DCSAY '           Delivered'  get lDelivered       EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo2 ;
		SAYCOLOR{|| IIF(lDelivered ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,GRA_CLR_PALEGRAY} )}

	@ 7,1 DCSAY '  Post Delivery Call' get oRecord:SENDLTR EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo2 ;
		SAYCOLOR{|| IIF(oRecord:SENDLTR='Y' ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,GRA_CLR_PALEGRAY} )}

	@ 4,40 DCPUSHBUTTONXP CAPTION 'Record Appointment ' ;
	 SIZE 20,2 ;
	 ACTION {|| _makeupappt(oRecord)};
	 CONFIG oConfig2  ;
	 BITMAP 8062 ;
	 PARENT oProgInfo2


	@ 32,0 DCTOOLBAR AoToolBar                               ;
		SIZE 140,2 BUTTONSIZE 12,2


@ 2,120 DCPUSHBUTTONXP CAPTION 'Record Be Back' ;
	 SIZE 20,2 ;
	 ACTION {|| _MAKEBEBACK(@oRecord),DC_GETREFRESH(GETLIST)} ;
	 CONFIG oConfig1  ;
	 BITMAP 7755      ;
	 PARENT oUpinfo

@ 4,120 DCPUSHBUTTONXP CAPTION 'Send Record to Vin ' ;
	 SIZE 20,2 ;
	 ACTION {|| NEWUPS->(DBGOTO(oRecord:Record_Number)),;
	 NEWUPS->(DC_DBGATHER(oRecord)),;
	 _senduptovin(NEWUPS->(RECNO()))};
	 CONFIG oConfig2  ;
	 BITMAP 7762 ;
	 PARENT oUpinfo


IF MANAGER()
	DCADDBUTTON CAPTION 'Payment;Calculator'                              ;
		ACTION {|| QUICK(0,0,0,0,0,0,0,0,0,0,.T.,0)  };
		PARENT AoToolBar      ;
		HIDE {|| IF(!MANAGER(),.T.,.F.)}
ENDIF

DCADDBUTTON CAPTION 'Comments '                              ;
	ACTION {|| UPNOTES()  };
	PARENT AoToolBar


DCADDBUTTON CAPTION 'Print; UPSheet '                              ;
	ACTION {|| NEWUPS->(DBGOTO(oRecord:Record_Number)),;
	 NEWUPS->(DC_DBGATHER(oRecord)),;
	PRINTUPSHEET(NEWUPS->(RECNO()))};
	PARENT AoToolBar

DCADDBUTTON CAPTION 'Desking;Tool '                              ;
		;//ACTION {|| NEWUPS->(DBGOTO(oRecord:Record_Number)),;
	   ;//NEWUPS->(DC_DBGATHER(oRecord)),;
		ACTION {|| UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)};
		PARENT AoToolBar


DCADDBUTTON CAPTION 'Buyers ;Order '                              ;
		ACTION {|| NEWUPS->(DBGOTO(oRecord:Record_Number)),;
	   NEWUPS->(DC_DBGATHER(oRecord)),;
		BuyersOrder(NEWUPS->(RECNO()),.f.,@oRecord),DC_GETREFRESH(GETLIST)};
		PARENT AoToolBar

IF MANAGER()
  DCADDBUTTON CAPTION 'Final Buyers; Order '                              ;
		SIZE 15                                              ;
		ACTION {|| NEWUPS->(DBGOTO(oRecord:Record_Number)),;
	   NEWUPS->(DC_DBGATHER(oRecord)),;
		BuyersOrder(NEWUPS->(RECNO()),.t.,@oRecord),DC_GETREFRESH(GETLIST)};
		PARENT AoToolBar
ENDIF

DCADDBUTTON CAPTION 'Appraisal; Form '                              ;         /// this will update Newups appraisal record
		ACTION {|| NEWUPS->(DBGOTO(oRecord:Record_Number)),;
	   NEWUPS->(DC_DBGATHER(oRecord)),;
		_appraiseform(NEWUPS->(RECNO()),oRecord:UPID,oRecord:SALESMAN,oRecord,.f.),;
		NEWUPS->(DBGOTO(oRecord:Record_Number)),;
	   NEWUPS->(DC_DBGATHER(oRecord))};
		PARENT AoToolBar

DCADDBUTTON CAPTION 'Review; eMails '                              ;
		ACTION {|| BROWEMAILLOG('',NEWUPS->EMAIL)} ;
		PARENT AoToolBar

DCADDBUTTON CAPTION 'Appointment; History '                              ;
	ACTION {||_chkapptsbrowse(NEWUPS->UpId)};
	PARENT AoToolBar



DCADDBUTTON CAPTION 'Exit '                              ;
	ACTION {||DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)};
	PARENT AoToolBar

DCREAD GUI TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA ;
		EVAL {|| SETAPPFOCUS(DC_GETOBJECT(GETLIST,'UPSTAT'))};
	 	FIT ADDBUTTONS OPTIONS GETOPTIONS;
	  TITLE 'Up Record Maintainance'
	  //	EVAL{|o| setappwindow(o)}

IF !XSTATUS
 RETURN NIL
ENDIF

NEWUPS->(DBGOTO(oRecord:Record_Number))
NEWUPS->(DC_DBGATHER(oRecord))

//DBCLEARSCOPE()
RETURN NIL





FUNCTION LOOKEMAILS()
	local cSender


	IF !UseDb({'PASSPROG'})
		 DC_WINALERT('Cannot Open Files     .')
	   	 RETURN FALSE
	ENDIF
	locate for PASSPROG->OPERATOR=USERINIT()                 ///  USERINIT() is a public var establised at login
	IF FOUND()
		cSender:=PAD(PASSPROG->email,50)
	endif
   BROWEMAILLOG(cSender,'')
RETURN NIL



static function _CHK4TYPECHANGE(nCurrUptype,oRecord)
	LOCAL GETLIST:={}, GETOPTIONS
	DCGETOPTIONS saywidth 0 autoresize sayfont '10.Courier New' getfont '10.Courier New'
	IF nCurrUpType > 4
		IF oRecord:UPTYPE < 5
			@ 10,0 DCSAY 'Please Verify the date this Up was converted and physically came in.' get oRecord:DATEIN
			DCREAD GUI MODAL ADDBUTTONS ENTEREXIT OPTIONS GETOPTIONS FIT TITLE 'Verify Date In' EVAL{|o| SETAPPWINDOW(o)}
			oRecord:ORIGDTIN:=oRecord:DATEIN
		ENDIF
	ENDIF
	RETURN TRUE

static function _chkbillstat(oRecord,getlist)
	IF !oRecord:Billstat $'ID'
		return FALSE
	ENDIF
	IF oRecord:Billstat='D'
		oRecord:DELIVERED:=.T.
	ELSE
		oRecord:DELIVERED:=.F.
	ENDIF
	DC_GETREFRESH(GETLIST)
	RETURN TRUE

static function _CHK4SHOWCLR(cUpId)
	LOCAL aColor
	SELECT UPAPPTS
	ORDSETFOCUS('UPAPPTS')
	SEEK  cUpId
	IF !FOUND()
		RETURN {GRA_CLR_BLACK,GRA_CLR_PALEGRAY}
	ENDIF
	IF FOUND()
		DO WHILE UPAPPTS->UPID=cUpid
			IF UPAPPTS->SHOWED
				aColor:= {GRA_CLR_BLACK,GRA_CLR_GREEN}
			else
				aColor:= {GRA_CLR_BLACK,GRA_CLR_RED}
			ENDIF
			IF UPAPPTS->APPTDATE >= DATE()
				IF UPAPPTS->SHOWED
				 aColor:= {GRA_CLR_BLACK,GRA_CLR_GREEN}
			   else
				 aColor:= {GRA_CLR_WHITE,GRA_CLR_DARKBLUE}
				ENDIF
			ENDIF
			SKIP alias UPAPPTS
		ENDDO
		SKIP-1 ALIAS UPAPPTS
	ENDIF
	RETURN aColor ///{GRA_CLR_BLACK,GRA_CLR_RED}

static function _CHK4SHOW(cUpId)
	select UPAPPTS
	ORDSETFOCUS('UPAPPTS')
	SEEK  cUpId
	IF !FOUND()
		RETURN FALSE
	ENDIF
	IF FOUND()
		DO WHILE UPAPPTS->UPID=cUpid
			IF UPAPPTS->SHOWED
				RETURN TRUE
			ENDIF
			SKIP alias UPAPPTS
		ENDDO
		SKIP-1 ALIAS UPAPPTS
	ENDIF
	RETURN FALSE



static function _CHK4APPTS(aAppts,aUPChars,aUpBoolean)
	select UPAPPTS
	ORDSETFOCUS('UPAPPTS')
	SEEK aUpChars[UPSC_CUPID]
	aSize(aAppts,0)
	IF !FOUND()
		RETURN FALSE
	ENDIF
	IF FOUND()
		DO WHILE UPAPPTS->UPID=aUpChars[UPSC_CUPID]
			AADD(aAppts,UPAPPTS->(RECNO()))
			aUpBoolean[UPSL_SHOWED]:=UPAPPTS->SHOWED
			SKIP alias UPAPPTS
		ENDDO
		SKIP-1 ALIAS UPAPPTS
		aUpBoolean[UPSL_APPOINT]:=.T.
	ENDIF
	RETURN TRUE

 static function _initArray(aArray,nSize,cnFill)
	ASize(aArray,0)
	ASize(aArray,nSize)                            // or whatever the size should be
	AFill(aArray,cnFill)
RETURN NIL

static Function _chkapptsbrowse(cUpId)
	local GETLIST:={},GETOPTIONS, aPres, oBrowseAppt ,aAppts:={} ,xStatus
	select UPAPPTS
	ORDSETFOCUS('UPAPPTS')
	UPAPPTS->(DBGOTOP())
	SEEK cUpId
	IF !FOUND()
		RETURN FALSE
	ENDIF
	IF FOUND()
		DO WHILE UPAPPTS->UPID=cUpId
			AADD(aAppts,UPAPPTS->(RECNO()))
			SKIP alias UPAPPTS
		ENDDO
		SKIP-1 ALIAS UPAPPTS
	ENDIF

	DC_SETSCOPEARRAY(aAppts)
	UPAPPTS->(DC_DBGOTOP())
	SEEK cUpId


	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, -1 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 10 }  }              /* Cell Height */

  	@ 5,5 DCSAY 'Appointment History'

	@ 8,5 DCBROWSE oBrowseAppt ALIAS 'UPAPPTS'                 ;
		SIZE 100,6                                     ;
		PRESENTATION aPres  ;
		FREEZELEFT {1,3} ;
		SCOPE;
		EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITDOWN

  DCBROWSECOL FIELD UPAPPTS->LASTNAME                     ;
		HEADER "Lastname" HCOLOR 1,5 PARENT oBrowseAppt   ;
		width 10
  DCBROWSECOL FIELD UPAPPTS->SM                     ;
		HEADER "SP" HCOLOR 1,5 PARENT oBrowseAppt   ;
		width 5


  DCBROWSECOL FIELD UPAPPTS->APPTDATE                     ;
		HEADER "Date" HCOLOR 1,5 PARENT oBrowseAppt   ;
		width 5

  DCBROWSECOL FIELD UPAPPTS->APPTTIME                     ;
		HEADER "TIME" HCOLOR 1,5 PARENT oBrowseAppt ;
		WIDTH 4
  DCBROWSECOL FIELD UPAPPTS->REASON                    ;
		HEADER "Reason" HCOLOR 1,5 PARENT oBrowseAppt


  DCBROWSECOL FIELD UPAPPTS->FIRSTAPPT                     ;
		HEADER "FirstAppt" HCOLOR 1,3 PARENT oBrowseAppt  ;
		EDITPROTECT{|| TRUE }

  DCBROWSECOL FIELD UPAPPTS->SHOWED                    ;
		HEADER "SHOWED" HCOLOR 1,5 PARENT oBrowseAppt COLOR{|| _UPSHOWCOLOR()}

  DCREAD GUI APPWINDOW MAINWINDOW():DRAWINGAREA TO XSTATUS FIT ADDBUTTONS OPTIONS GETOPTIONS TITLE 'Appointment Browse' //EVAL{|o| setappwindow()}

  IF !xStatus
	RETURN TRUE
  ENDIF
  RETURN TRUE




static function _MAKEUPAPPT(oRecord)
  LOCAL GETLIST:={},GETOPTIONS, dApptDate:=CTOD('  /  /  '), cReason:=space(120), cTime:=SPACE(10), xStatus ,lFirst:=.t., lShowed:=.f.
  DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
  dApptdate:=DC_GUIPOPDATE(,,,@xStatus,'Make Appointment for.. '+alltrim(oRecord:FIRSTNAME)+' '+alltrim(oRecord:LASTNAME),.t.,1)

  IF oRecord:DATEIN <> NEWUPS->BEBACK3         /// BEBACK3 IS ACTUALLY THE ORIGINAL DATE IN FIELD
	  lFirst:=.f.
  ENDIF

  @ 10,0 DCSAY 'First Name' get oRecord:FIRSTNAME EDITPROTECT{|| TRUE }
  @ 11,0 DCSAY 'Last Name ' Get oRecord:LASTNAME  EDITPROTECT{|| TRUE }
  @ 13,0 DCSAY 'Enter The Date of Next Appointment  ' get dApptdate
  @ 14,0 DCSAY 'Enter Time of Appointment           ' get cTime
  @ 15,0 DCSAY 'Reason For Appointment'
  @ 16,0	DCMULTILINE cReason FONT "10.Courier" SIZE 60,2  NOHORIZSCROLL NOVERTSCROLL

  @ 18,0 DCCHECKBOX lFirst PROMPT 'Check here if this is the FIRST Appointment.'
  @ 20,0 DCCHECKBOX lShowed PROMPT 'Check here if this Up already showed up.   '


  DCREAD GUI MODAL ADDBUTTONS ENTEREXIT OPTIONS GETOPTIONS FIT TITLE 'Set Appointment  ' EVAL{|o| SETAPPWINDOW(o)}
  IF !xStatus
	RETURN NIL
  ENDIF
  IF UPAPPTS->(DC_ADDREC())
	REPLACE UPAPPTS->UPID WITH oRecord:UPID
	REPLACE UPAPPTS->APPTDATE WITH dApptDate
	REPLACE UPAPPTS->LASTNAME WITH oRecord:LASTNAME
	REPLACE UPAPPTS->FIRSTNAME WITH oRecord:FIRSTNAME
	REPLACE UPAPPTS->REASON WITH cReason
	REPLACE UPAPPTS->APPTTIME with cTime
	REPLACE UPAPPTS->PHONE WITH oRecord:AREA+oRecord:UPID
	REPLACE UPAPPTS->EMAIL WITH oRecord:EMAIL
	REPLACE UPAPPTS->SM WITH oRecord:SALESMAN
	REPLACE UPAPPTS->ORIGDATE WITH oRecord:BEBACK3
	REPLACE UPAPPTS->INTEREST WITH oRecord:INTEREST
	REPLACE UPAPPTS->LASTDATE WITH oRecord:DATEIN
	REPLACE UPAPPTS->NEWUSED WITH oRecord:NEWUSED
	REPLACE UPAPPTS->UPSOURCE WITH oRecord:ORIGSOURCE
	REPLACE UPAPPTS->UPTYPE WITH oRecord:UPTYPE
	IF lFirst
		REPLACE UPAPPTS->FIRSTAPPT WITH TRUE
		REPLACE UPAPPTS->FOLLOWAPPT WITH FALSE
	ELSE
		REPLACE UPAPPTS->FIRSTAPPT WITH FALSE
		REPLACE UPAPPTS->FOLLOWAPPT WITH TRUE
	ENDIF
	IF lShowed
		REPLACE UPAPPTS->SHOWED WITH TRUE
		oRecord:appoint:=.t.
	ENDIF
  ENDIF

  RETURN NIL

 function UpApptBrowse(lOpenfiles,cSM)
	LOCAL GETLIST:={}, getoptions, xStatus , aPres, oBrowse, oToolbar, nRec:=0
	DEFAULT cSM:=space(2)
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '8.Courier New' GETFONT '8.Courier New' AUTORESIZE
	DEFAULT lOpenFiles:=.f.
	IF lOpenfiles
		IF !OPENUPSFILES()
		  RETURN NIL
		ENDIF
	ENDIF

	IF !EMPTY(cSM)
	  SELECT UPAPPTS
	  ORDSETFOCUS('UPAPSMDT')
	  DBGOTOP()
	  SET FILTER TO UPAPPTS->SM=cSm
	  SEEK cSM+DTOS(DATE()-1)
	 ELSE
	 SELECT UPAPPTS
	 ORDSETFOCUS('UPAPPTDT')
	 SEEK DATE()-1
	ENDIF
	/*
	DC_SETSCOPE(0,DATE()-1)
	DC_SETSCOPE(1,DATE()) */


	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

	@ 1,1 DCTOOLBAR oToolBar                               ;
		SIZE 115, 1.5

	DCADDBUTTON CAPTION '1.Send eMail '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("1")};
	ACTION {||upApptemail()};
	PARENT oToolBar

	DCADDBUTTON CAPTION '2.Screen By ; SalesPerson '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("2")};
	ACTION {|| UPAPPTSCREENsm(),oBrowse:REFRESHALL(),oBrowse:FORCESTABLE()};
	PARENT oToolBar

	DCADDBUTTON CAPTION '3.Screen By ; Date '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("3")};
	ACTION {|| UPAPPTSCREENDT(),oBrowse:REFRESHALL(),oBrowse:FORCESTABLE()};
	PARENT oToolBar

	DCADDBUTTON CAPTION '4.Screen For ; No Shows '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("4")};
	ACTION {|| UPAPPTSCREENNS(),oBrowse:REFRESHALL(),oBrowse:FORCESTABLE()};
	PARENT oToolBar


	DCADDBUTTON CAPTION '5.Appointment ; Reports '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("5")};
	ACTION {|| nRec:=UPAPPTS->(RECNO()),_DailyApptRep(),UPAPPTS->(DBGOTO(nRec)),oBrowse:REFRESHALL(),oBrowse:FORCESTABLE()};
	PARENT oToolBar

	@ 4,1 DCBROWSE oBrowse ALIAS 'UPAPPTS'                 ;
		SIZE 130,20                                     ;
		PRESENTATION aPres  ;
		FREEZELEFT {1,2,3} ;
		SORTSCOLOR 0,2 ;
		SORTUCOLOR 1,6   ;
		EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITACROSS


  DCBROWSECOL FIELD UPAPPTS->APPTDATE                     ;
		HEADER "**ApptDate" HCOLOR 1,5 PARENT oBrowse   ;
		SORT{|| UPAPPTS->(ORDSETFOCUS('UPAPPTDT')),dbgotop(),oBrowse:REFRESHALL()}

  DCBROWSECOL FIELD UPAPPTS->SM                     ;
		HEADER "*SM" HCOLOR 1,3 PARENT oBrowse   ;
		SORT{|| UPAPPTS->(ORDSETFOCUS('UPAPPTSM')),dbgotop(),oBrowse:REFRESHALL()};
	   EDITPROTECT{|| TRUE }

  DCBROWSECOL FIELD UPAPPTS->LASTNAME                     ;
		HEADER "*LastName" HCOLOR 1,5 PARENT oBrowse   ;
		SORT{|| UPAPPTS->(ORDSETFOCUS('UPAPPTNM')),dbgotop(),oBrowse:REFRESHALL()} ;
		WIDTH 10


  DCBROWSECOL FIELD UPAPPTS->FIRSTNAME                     ;
		HEADER "FIRST" HCOLOR 1,5 PARENT oBrowse   ;
		WIDTH 5

  DCBROWSECOL FIELD UPAPPTS->ORIGDATE                     ;
		HEADER "OrigDateIn" HCOLOR 1,3 PARENT oBrowse ;
		EDITPROTECT{|| TRUE }  ;
		WIDTH 8

  DCBROWSECOL FIELD UPAPPTS->LASTDATE                     ;
		HEADER "LastDateIn" HCOLOR 1,3 PARENT oBrowse ;
		EDITPROTECT{|| TRUE }  ;
		WIDTH 8


  DCBROWSECOL FIELD UPAPPTS->FIRSTAPPT                     ;
		HEADER "FirstAppt" HCOLOR 1,3 PARENT oBrowse  ;
		EDITPROTECT{|| TRUE }

  DCBROWSECOL FIELD UPAPPTS->FOLLOWAPPT                     ;
		HEADER "FollowUpAppt" HCOLOR 1,3 PARENT oBrowse ;
		EDITPROTECT{|| TRUE }

  DCBROWSECOL FIELD UPAPPTS->SHOWED                    ;
		HEADER "SHOWED" HCOLOR 1,5 PARENT oBrowse COLOR{|| _UPSHOWCOLOR()}

  DCBROWSECOL FIELD UPAPPTS->APPTTIME                     ;
		HEADER "TIME" HCOLOR 1,5 PARENT oBrowse ;
		WIDTH 4

  DCBROWSECOL FIELD UPAPPTS->UPID                     ;
		HEADER "**UpId" HCOLOR 1,3 PARENT oBrowse   ;
		SORT{|| UPAPPTS->(ORDSETFOCUS('UPAPPT')),dbgotop(),oBrowse:REFRESHALL()};
	   EDITPROTECT{|| TRUE }    ;
		WIDTH 6

  DCBROWSECOL FIELD UPAPPTS->REASON                     ;
		HEADER "Reason" HCOLOR 1,5 PARENT oBrowse ;
		WIDTH 10


  DCBROWSECOL FIELD UPAPPTS->PHONE                     ;
		HEADER "PHONE" HCOLOR 1,5 PARENT oBrowse ;
		WIDTH 10

  DCBROWSECOL FIELD UPAPPTS->EMAIL                     ;
		HEADER "eMail" HCOLOR 1,5 PARENT oBrowse ;
		WIDTH 10



    DCBROWSECOL FIELD UPAPPTS->UPSOURCE                     ;
		HEADER "UpSource" HCOLOR 1,3 PARENT oBrowse ;
		EDITPROTECT{|| TRUE }

  DCBROWSECOL FIELD UPAPPTS->UPTYPE                     ;
		HEADER "UpType" HCOLOR 1,3 PARENT oBrowse ;
		EDITPROTECT{|| TRUE }

  DCBROWSECOL FIELD UPAPPTS->NEWUSED                     ;
		HEADER "NewUsed" HCOLOR 1,5 PARENT oBrowse ;

  DCBROWSECOL FIELD UPAPPTS->INTEREST                     ;
		HEADER "Interest" HCOLOR 1,3 PARENT oBrowse

  DCREAD GUI ;
		FIT ;
		TITLE 'Up Appointment Browser' ;
		OPTIONS GETOPTIONS ;
		ADDBUTTONS


  RETURN NIL

static function _DailyApptRep()
	LOCAL GETLIST:={}, GETOPTIONS, xStatus, cSm:=SPACE(2), dDate1:=DATE(),dDate2:=CTOD('  /  /  '),lShowed:=.F., nPage:=0
	local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New',cPhone
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
	TOP_MAR(2)
	BOT_MAR(2)
	PAGE_LEN(62)

	@ 5,0 DCSAY 'Appointment Printing Options'
	@ 7,0 DCSAY 'Enter Beginning Date' get dDate1
	@ 8,0 DCSAY 'Enter Ending Date   ' get dDate2
	@ 9,0 DCSAY 'Enter Sales Person initials or leave blank for all.'  get cSM picture '@!'
	@ 11,0 DCCHECKBOX lShowed PROMPT 'Check here to ONLY include NO Shows.'
   DCREAD GUI MODAL ADDBUTTONS ENTEREXIT OPTIONS GETOPTIONS FIT TITLE 'Appointment Reports' EVAL{|o| SETAPPWINDOW(o)}

	SELECT UPAPPTS
	ORDSETFOCUS('UPAPPTDT')
	SEEK dDate1
	IF !FOUND()
		DC_WINALERT('No Appointments for that date!')
		RETURN NIL
	ENDIF


	IF !empty(cSm)
		IF lShowed
		 UPAPPTS->(DBSETFILTER({|| UPAPPTS->SM==cSm .AND.UPAPPTS->SHOWED=.F. }))
   	else
		UPAPPTS->(DBSETFILTER({|| UPAPPTS->SM==cSm}))
	   ENDIF
	ENDIF

	IF empty(cSm)
	 IF lShowed
		UPAPPTS->(DBSETFILTER({|| UPAPPTS->SHOWED=.F.}))
	 ENDIF
	ENDIF

	IF !Print_Choice('Daily Appointments', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
   ENDIF
   IF !PrintOn('Daily Appointments', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
   ENDIF
	_UPAPPTHDR(@nPage,dDate1,dDate2)
	skipaline()
	DO WHILE !UPAPPTS->(EOF())
		IF UPAPPTS->ApptDate >=dDate1
			IF UPAPPTS->ApptDate <= DdATE2
				cPhone:=_formatphn(UPAPPTS->PHONE)
				@ DC_PRINTERROW()+1,0 DCPRINT SAY UPAPPTS->ApptDate
				@ DC_PRINTERROW(),10 DCPRINT SAY UPAPPTS->ApptTime
				@ DC_PRINTERROW(),20 DCPRINT SAY UPAPPTS->Sm
				@ DC_PRINTERROW(),25 DCPRINT SAY substr(UPAPPTS->LastName,1,19) FONT '8.Courier New Bold'
				@ DC_PRINTERROW(),45 DCPRINT SAY substr(UPAPPTS->FirstName,1,14)

				IF !UPAPPTS->SHOWED
				 IF UPAPPTS->ApptDate >= DATE()
				  @ DC_PRINTERROW(),60 DCPRINT SAY 'PENDING' FONT '10.Courier Bold'
				 else
				  @ DC_PRINTERROW(),60 DCPRINT SAY 'NO SHOW' FONT '10.Courier Bold'
				 ENDIF
				ENDIF

				IF UPAPPTS->SHOWED
					@ DC_PRINTERROW(),60 DCPRINT SAY 'Yes'
				ENDIF
				IF !UPAPPTS->FirstAppt
	   			@ DC_PRINTERROW(),70 DCPRINT SAY 'FirstTime'
				else
					@ DC_PRINTERROW(),70 DCPRINT SAY 'Followup'
				ENDIF

				@ DC_PRINTERROW(),80 DCPRINT SAY UPAPPTS->uPSource
				@ DC_PRINTERROW(),85 DCPRINT SAY UPAPPTS->UPType
				@ DC_PRINTERROW(),90 DCPRINT SAY UPAPPTS->Interest
				@ DC_PRINTERROW(),105 DCPRINT SAY UPAPPTS->OrigDate
				@ DC_PRINTERROW(),115 DCPRINT SAY cPhone
				@ DC_PRINTERROW(),132 DCPRINT SAY UPAPPTS->NewUsed
				@ DC_PRINTERROW()+1,0 DCPRINT SAY UPAPPTS->eMail
				@ DC_PRINTERROW(),50 DCPRINT SAY  UPAPPTS->Reason
				IF DCPAGEEJECT()
					_UPAPPTHDR(@nPage,dDate1,dDate2)
				ENDIF
				SKIPALINE()
				IF DCPAGEEJECT()
					_UPAPPTHDR(@nPage,dDate1,dDate2)
				ENDIF
			ENDIF  //DDATE2
		ENDIF     /// DDATE1
		SKIP ALIAS UPAPPTS
	ENDDO
	PRINTOFF(oPrinter)
	UPAPPTS->(DBCLEARFILTER())

	RETURN NIL


static function _UPAPPTHDR(nPage,dDate1,dDate2)
	nPage++
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Appointment Reports '+ dtoc(dDate1)+' to '+dtoc(dDate1)+ ' Page '+alltrim(str(nPage))
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'ApptDate' font '8.Courier New Bold'
	@ DC_PRINTERROW(),10 DCPRINT SAY 'Time'      font '8.Courier New Bold'
	@ DC_PRINTERROW(),20 DCPRINT SAY 'SP'        font '8.Courier New Bold'
	@ DC_PRINTERROW(),25 DCPRINT SAY 'LastName'  font '8.Courier New Bold'
	@ DC_PRINTERROW(),45 DCPRINT SAY 'FirstName' font '8.Courier New Bold'
	@ DC_PRINTERROW(),60 DCPRINT SAY 'Showed'    font '8.Courier New Bold'
	@ DC_PRINTERROW(),70 DCPRINT SAY 'FirstAppt' font '8.Courier New Bold'
	@ DC_PRINTERROW(),80 DCPRINT SAY 'Srce'    font '8.Courier New Bold'
	@ DC_PRINTERROW(),85 DCPRINT SAY 'UpType'    font '8.Courier New Bold'
	@ DC_PRINTERROW(),92 DCPRINT SAY 'Interest' font '8.Courier New Bold'
	@ DC_PRINTERROW(),105 DCPRINT SAY 'OrigDateIn' font '8.Courier New Bold'
	@ DC_PRINTERROW(),115 DCPRINT SAY 'Phone'      font '8.Courier New Bold'
	@ DC_PRINTERROW(),129 DCPRINT SAY 'N/U'        font '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'eMail'      font '8.Courier New Bold'
	@ DC_PRINTERROW(),50 DCPRINT SAY 'Reason For Appointment ' font '8.Courier New Bold'
	RETURN NIL

static function _UPSHOWCOLOR()
  IF UPAPPTS->SHOWED
	 RETURN {GRA_CLR_BLACK,GRA_CLR_GREEN}
  ENDIF
  IF !UPAPPTS->SHOWED
	IF UPAPPTS->APPTDATE >=DATE()
	 RETURN {GRA_CLR_WHITE,GRA_CLR_DARKBLUE}
	ENDIF
	IF UPAPPTS->APPTDATE < DATE()
	 RETURN {GRA_CLR_BLACK,GRA_CLR_RED}
	ENDIF
  ENDIF
  RETURN {GRA_CLR_BLACK,GRA_CLR_RED}

static function UPAPPTSCREENsm()
	LOCAL cSM:=space(2), dDate1:=DATE(),dDate2:=DATE()
	LOCAL GETLIST:={}, GETOPTIONS
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
	@ 10,0 DCSAY 'Enter Sales Person Initials to Screen For ' get cSm
	@ 11,0 DCSAY 'Enter Beginning Date                      ' get dDate1
	DCREAD GUI MODAL ADDBUTTONS ENTEREXIT OPTIONS GETOPTIONS FIT TITLE 'SM Date Screen' EVAL{|o| SETAPPWINDOW(o)}

	SELECT UPAPPTS
	ORDSETFOCUS('UPAPSMDT')
	SET FILTER TO UPAPPTS->SM=cSm
	UPAPPTS->(DBGOTOP())
	SEEK cSm+DTOS(dDate1)
	return nil

static function UPAPPTSCREENNS()
	LOCAL cSM:=space(2), dDate1:=DATE(),dDate2:=DATE()
	LOCAL GETLIST:={}, GETOPTIONS
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
	@ 11,0 DCSAY 'Enter Beginning Date                      ' get dDate1
	DCREAD GUI MODAL ADDBUTTONS ENTEREXIT OPTIONS GETOPTIONS FIT TITLE 'SM Date Screen' EVAL{|o| SETAPPWINDOW(o)}

	SELECT UPAPPTS
	ORDSETFOCUS('UPAPPTDT')
	SET FILTER TO UPAPPTS->SHOWED=.F.
	UPAPPTS->(DBGOTOP())
	SEEK dDate1
	return nil



static function UPAPPTSCREENdt()
	LOCAL  dDate1:=DATE()-1,dDate2:=DATE()
	LOCAL GETLIST:={}, GETOPTIONS
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
	@ 11,0 DCSAY 'Enter Beginning Date                      ' get dDate1
	@ 12,0 DCSAY 'Enter Ending Date                         ' get dDate2

	DCREAD GUI MODAL ADDBUTTONS ENTEREXIT OPTIONS GETOPTIONS FIT TITLE 'Appts Date Screen' EVAL{|o| SETAPPWINDOW(o)}
	DbCLEARSCOPE()
	SELECT UPAPPTS
	ORDSETFOCUS('UPAPPTDT')
	SEEK dDate1
	DC_SETSCOPE(0,dDate1)
	DC_SETSCOPE(1,dDate2)
	dbgotop()

	return nil



static function FIRSTCHECK()
	local getlist:={}
	@ nil,nil DCCHECKBOX UPAPPTS->FIRSTAPPT
return TRUE

static function FOLLOWCHECK()
	local getlist:={}
	@ nil,nil DCCHECKBOX UPAPPTS->FOLLOWAPPT
	return TRUE


static function _MAKEBEBACK(oRecord)
	LOCAL GETLIST:={},GETOPTIONS
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE

	IF oRecord:UPTYPE > 2
		DC_WINALERT('You cannot record the be back of an up UNLESS it is already a 1 or 2,;Ups that are entered as 5 or six must be converted to 1 or 2.;Ups that are entered as 3 or 4 will remain that way.')
		RETURN NIL
	ENDIF

	IF empty(oRecord:DBEBACK1)
		oRecord:BEBACK1:=DATE()
		@ 5,0 DCSAY 'Record First Be Back Date of' get oRecord:BEBACK1
	  else
		oRecord:BEBACK2:=DATE()
		@ 5,0 DCSAY 'Record Second Be Back Date of' get oRecord:BEBACK2
	endif
	DCREAD GUI MODAL ADDBUTTONS ENTEREXIT OPTIONS GETOPTIONS FIT TITLE 'Be Back s' EVAL{|o| SETAPPWINDOW(o)}
	return nil


static function _appraiseform(nRec,cUpid,cSM,oDeskRecord,lFromBrowse)
	local nCopies:=1,nOrient:=1,nDefmode:=4, oPrinter,cOutfile:='',oStatic,oDisclose2,oOwner,oExtcon,oTradein, i
	local clSearchService:='Y', cSrchnm:=UPPER(NEWUPS->LASTNAME),oRecord ,lNewAppraisal:=.f.
	local getlist:={}, cColor:=space(20)
	local getoptions, xStatus, nServrec:=0, nAcv:=0 ,cManagerInit:=space(3) , nLines:=0, dDate:=date(),oVinOpts
	local oStdopts,oExtras,oSafety,oMedia,oEngine,oOthers,oDisclose,oBodywork,oAirbagdep,oOdomacc,oFramedam,oRepreq,oAfterperf
	DEFAULT lFromBrowse:=.f.
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE

  IF !lFromBrowse
  	SELECT UCAPPRSL
  	ORDSETFOCUS('UCAPPVIN')
  	UCAPPRSL->(DBGOTOP())
  	SEEK oDeskRecord:TRADEVIN
  	IF found()
  		lNewAppraisal:=.f.
  		oRecord:=DC_DBRECORD():NEW()
  		UCAPPRSL->(DC_DBSCATTER(oRecord))

  	endif
  	IF !found()
  		  lNewAppraisal:=.t.
  		  oRecord:=DC_DBRECORD():NEW()
  		  UCAPPRSL->(DB_INIT(oRecord))

  		  @ 10,0 DCSAY 'Search Service Base for this Vehicle Y/N' get clSearchService PICTURE '@! L'
  		  @ 13,0 DCSAY 'Enter Search Last Name Up To 20 Characters Esc=exit'  GET cSRCHNM PICTURE "@!";
  		 	 WHEN {|| IIF(clSearchService='Y',.t. ,.f. )}
  		  dcread gui MODAL enterexit to xstatus options getoptions fit title 'Search Customer Base'   ;
  			EVAl{|o| setappwindow(o)}
  		  if !xstatus
  			return xStatus
  		  endif

  	  IF clSearchService='Y'
  			@ 3,0 DCSAY 'Enter Search Last Name Up To 20 Characters Esc=exit'  GET cSRCHNM PICTURE "!!!!!!!!!!!!!!!!!!!!"   ;
  				GETTOOLTIP('YOU MAY ENTER UP TO 20 CHARACTERS, HOWEVER 3 OR 4 IS USUALLY ENOUGH ;TO PLACE THE FILE POINTER AT THE APPROXIMATE POINT TO EASILY ISOLATE THE RECORD NEEDED.;  YOU MAY PRESS ESCAPE TO EXIT ')
  			DCREAD GUI modal BUTTONS 2 ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'SEARCH FOR NAME  ' ;
  				eval{|o| setappwindow(o)}
  			SELECT SERVCUST
  			SET SOFTSEEK ON
  			LOOKSERV(@nServREC,cSRCHNM)
  			SET SOFTSEEK OFF
  		 ENDIF  &    &ACON
  		 IF nServRec <> 0
  			GOTO nServRec                                  && IF ONE IS SELECTED
  			oRecord:Vin:=SERVCUST->vin
  			oRecord:Year:=SERVCUST->YEAR
  			oRecord:Make:=SERVCUST->make
  			oRecord:Model:=SERVCUST->cardesc
  			cColor:=SERVCUST->color
  		 else                         /// if one is not selected  get data from oDeskrecord
  			IF VALTYPE(oDeskRecord)='O'         /// if oDeskrecord is an object
  		    oRecord:Year:=oDeskRecord:TRADEYR
  		    oRecord:Model:=oDeskRecord:TRADEDESC
  		    oRecord:Make:= oDeskRecord:TRADEMAKE
  		    oRecord:TradeAcv:=oDeskRecord:TRADEACV
  		    oRecord:Vin:=oDeskRecord:TRADEVIN
  		    cColor:=oDeskRecord:TRADECOLOR
  		    oRecord:Mileage:=oDeskRecord:TRADEMILES
  		   ENDIF
  		 ENDIF
  		 IF clSearchService='N'
  			IF VALTYPE(oDeskRecord)='O'         /// if oDeskrecord is an object
  		    oRecord:Year:=oDeskRecord:TRADEYR
  		    oRecord:Model:=oDeskRecord:TRADEDESC
  		    oRecord:Make:= oDeskRecord:TRADEMAKE
  		    oRecord:TradeAcv:=oDeskRecord:TRADEACV
  		    oRecord:Vin:=oDeskRecord:TRADEVIN
  		    cColor:=oDeskRecord:TRADECOLOR
  		    oRecord:Mileage:=oDeskRecord:TRADEMILES
  		   ENDIF
  		 ENDIF

  		 oRecord:APPRDATE := dDate
  		 oRecord:UPID:= cUpid
  		 oRecord:SM:= cSm
  		 oRecord:LASTNAME := NEWUPS->LASTNAME
  		 oRecord:FIRSTNAME := NEWUPS->Firstname
  	ENDIF
ENDIF  //if !lfrombrowse

IF lFromBrowse    /// if called from browser create rec object
	SELECT UCAPPRSL
	oRecord:=DC_DBRECORD():NEW()
  	UCAPPRSL->(DC_DBSCATTER(oRecord))
ENDIF

cColor:=UCAPPRSL->COLOR


	@ 0,0 DCSTATIC TYPE XBPSTATIC_TYPE_RECESSEDBOX ;
			OBJECT oStatic SIZE 132,32
	@ 1,1 DCSAY '        Vin ' get oRecord:Vin               PARENT oStatic PICTURE '@!' VALID{|| APPRTESTVIN(oRecord:Vin,oRecord:Year)}
	@ 2,1 DCSAY '        Year' get oRecord:Year PARENT oStatic
	@ 3,1 DCSAY '        Make' get oRecord:Make       PARENT oStatic  PICTURE '@!'
	@ 4,1 DCSAY '       Model' get oRecord:Model      PARENT oStatic  PICTURE '@!'
	@ 1,41 DCSAY '  TrimLevel' get oRecord:series     PARENT oStatic  PICTURE '@!'
		@ 2,41 DCSAY '   Ext Color' get cColor      PARENT oStatic  PICTURE '@!'
	@ 3,41 DCSAY '    Interior' get oRecord:Interior     PARENT oStatic
	@ 4,41 DCSAY '     Mileage' get oRecord:Mileage picture '999999' PARENT oStatic
	@ 4,80 DCSAY 'Number Doors' get oRecord:Numdoors picture '9' PARENT oStatic
	@ 1,100 DCSAY oRecord:LastName  PARENT oStatic
	@ 2,100 DCSAY oRecord:Salesman PARENT oStatic



	@ 3,100 DCPUSHBUTTON CAPTION 'Vin Decoder' PARENT oStatic;
			CARGO 'CANCEL' ;
		 	ACTION {|| VINDECODER(oRecord:Vin,@oRecord),DC_GETREFRESH(GETLIST)} SIZE 20,2


	@ 5.5,1 DCGROUP oEngine CAPTION  'PowerTrain' SIZE 40,7.2 PARENT oStatic
		@ 1,1 DCSAY ' Engine Size Liters' get oRecord:Engsize PARENT oEngine
		@ 2,1 DCSAY '   Number Cylinders' get oRecord:Cylinders PARENT oEngine
		@ 3,1 DCSAY '   Transmission A/M' get oRecord:Trans PARENT oEngine  PICTURE '@!'
		@ 4,1 DCSAY 'Transmission Speeds' get oRecord:Speeds PARENT oEngine  PICTURE '9'
		@ 5,1 DCCHECKBOX oRecord:AWD PROMPT 'All Wheel or 4 Wheel Drive' PARENT oEngine  TABSTOP
		@ 6,1 DCCHECKBOX oRecord:HYBRID PROMPT 'Hybrid Engine' PARENT oEngine  TABSTOP

	@ 13,1 DCGROUP oSafety CAPTION 'Safety Features' SIZE 40,5 PARENT oStatic   TABSTOP
		@ 1,1 DCCHECKBOX oRecord:Abs PROMPT 'Anti Lock Brakes        ' PARENT oSafety   TABSTOP
		@ 2,1 DCCHECKBOX oRecord:Sidebags PROMPT 'Side Impact AirBags' PARENT oSafety   TABSTOP
		@ 3,1 DCCHECKBOX oRecord:OnStar PROMPT 'GM Onstar            ' PARENT oSafety   TABSTOP

	@ 18,1 DCGROUP oStdOpts CAPTION 'Standard Equip' SIZE 40,6 PARENT oStatic   TABSTOP
		@ 1,1 DCCHECKBOX oRecord:CD PROMPT 'CD Player                ' PARENT oStdOpts  TABSTOP
		@ 2,1 DCCHECKBOX oRecord:locks PROMPT 'Power Locks           ' PARENT oStdOpts  TABSTOP
		@ 3,1 DCCHECKBOX oRecord:Windows PROMPT 'Power Windows       ' PARENT oStdOpts  TABSTOP
		@ 4,1 DCCHECKBOX oRecord:AC PROMPT 'Air Conditioning         ' PARENT oStdOpts  TABSTOP


	@ 5.5,41 DCGROUP oExtras CAPTION 'Media/Luxury UpGrades' SIZE 40,12 PARENT oStatic
		@ 1,1 DCCHECKBOX oRecord:Dvd PROMPT 'DVD Player              ' PARENT oExtras  TABSTOP
		@ 2,1 DCCHECKBOX oRecord:SatRad PROMPT 'Satellite XM Sirius  ' PARENT oExtras  TABSTOP
		@ 3,1 DCCHECKBOX oRecord:MP3 PROMPT 'MP3 Player              ' PARENT oExtras  TABSTOP
		@ 4,1 DCCHECKBOX oRecord:Nav PROMPT 'Navigation System       ' PARENT oExtras  TABSTOP
	  	@ 5,1 DCCHECKBOX oRecord:PowSeat PROMPT 'Power Driver Seat   ' PARENT oExtras  TABSTOP
		@ 6,1 DCCHECKBOX oRecord:HeatSeat PROMPT 'Heated Seats       ' PARENT oExtras  TABSTOP
		@ 7,1 DCCHECKBOX oRecord:Leather PROMPT 'Leather Interior    ' PARENT oExtras  TABSTOP
		@ 8,1 DCCHECKBOX oRecord:Sunroof PROMPT 'Sun Roof            ' PARENT oExtras  TABSTOP
		@ 9,1 DCCHECKBOX oRecord:BigWheels PROMPT '20 Inch Wheels    ' PARENT oExtras  TABSTOP
		@ 10,1 DCCHECKBOX oRecord:Rearcam PROMPT 'Rear Camera Monitor ' PARENT oExtras  TABSTOP

	@ 5.5,82 DCGROUP oVinOpts CAPTION 'Decoded Options' SIZE 50,11  parent oStatic
		@ 1,1 DCMULTILINE oRecord:VINOPTIONS  SIZE 40,10   EDITPROTECT{||.T.}    ;
       font '10.Arial' ;
		 PARENT oVinOpts;
		 COLOR 1,BD_LEMONCREME
	@ 8,41  DCPUSHBUTTON PROMPT 'Print' SIZE 9,2;
	ACTION {|| PRNTVEHCOMMENTS(oRecord:VINOPTIONS)}   PARENT oVinopts



	@ 18,41 DCGROUP oOthers CAPTION 'Other Options' SIZE 47,6 PARENT oStatic   TABSTOP
		@ 1,1 DCSAY 'OptCode' get oRecord:FEAT1 VALID {|| getfeatinfo(@oRecord:Feat1,@oRecord:Other1),DC_GETREFRESH(GETLIST)} PARENT  oOthers TABSTOP
		@ 1,15 DCSAY 'Option' get oRecord:Other1 PARENT  oOthers TABSTOP

		@ 2,1 DCSAY 'OptCode' get oRecord:FEAT2 VALID {|| getfeatinfo(@oRecord:Feat2,@oRecord:Other2),DC_GETREFRESH(GETLIST)} PARENT  oOthers TABSTOP
		@ 2,15 DCSAY 'Option' get oRecord:Other2 PARENT  oOthers TABSTOP

		@ 3,1 DCSAY 'OptCode' get oRecord:FEAT3 VALID {|| getfeatinfo(@oRecord:Feat3,@oRecord:Other3),DC_GETREFRESH(GETLIST)} PARENT  oOthers TABSTOP
		@ 3,15 DCSAY 'Option' get oRecord:Other3 PARENT  oOthers TABSTOP
		@ 4,1 DCSAY 'OptCode' get oRecord:FEAT4 VALID {|| getfeatinfo(@oRecord:Feat4,@oRecord:Other4),DC_GETREFRESH(GETLIST)} PARENT  oOthers TABSTOP
		@ 4,15 DCSAY 'Option' get oRecord:Other4 PARENT  oOthers TABSTOP
		@ 5,1 DCSAY 'OptCode' get oRecord:FEAT5 VALID {|| getfeatinfo(@oRecord:Feat5,@oRecord:Other5),DC_GETREFRESH(GETLIST)} PARENT  oOthers TABSTOP
		@ 5,15 DCSAY 'Option' get oRecord:Other5 PARENT  oOthers TABSTOP


	@ 18,88 DCGROUP oDisclose2 CAPTION 'Vehicle Disclosures' SIZE 48,4 PARENT oStatic

		 @ 1,1 DCGROUP oOwner CAPTION 'First Owner ' SIZE 45,1 PARENT oDisclose2
		 @ 0,30 DCRADIO oRecord:Owner VALUE 'Y'  PROMPT 'YES' COLOR 1,BD_KEYLIME ;
       PARENT oOwner
		 @ 0,38 DCRADIO oRecord:Owner VALUE 'N' PROMPT 'No' COLOR 1,BD_SALMON ;
       PARENT oOwner

		 @ 2,1 DCGROUP oExtcon CAPTION 'Extended Service Contract' SIZE 45,1 PARENT oDisclose2
		 @ 0,30 DCRADIO oRecord:Extcon VALUE 'Y'  PROMPT 'YES' COLOR 1,BD_KEYLIME ;
       PARENT oExtcon
		 @ 0,38 DCRADIO oRecord:Extcon VALUE 'N' PROMPT 'No' COLOR 1,BD_SALMON ;
       PARENT oExtcon

		 @ 3,1 DCGROUP oTradein CAPTION 'Trading for In Stock ' SIZE 45,1 PARENT oDisclose2
		 @ 0,30 DCRADIO oRecord:Instock VALUE 'Y'  PROMPT 'YES' COLOR 1,BD_KEYLIME ;
       PARENT oTradein
		 @ 0,38 DCRADIO oRecord:InStock VALUE 'N' PROMPT 'No' COLOR 1,BD_SALMON ;
       PARENT oTradein

	 @ 23,90 DCSAY 'Current Payments' GET oRecord:CurrPay picture '99999.99'  PARENT oStatic




	@ 24,51 DCSAY 'Comment' get oRecord:Comments   PARENT oStatic

	@ 26,1 DCGROUP oDisclose CAPTION 'Vehicle Disclosures' SIZE 80,6 PARENT oStatic
		 @ 1,1 DCGROUP oAirbagDep CAPTION 'Has The Air Bag Ever Been Deployed ' SIZE 80,1 PARENT oDisclose
		 @ 0,50 DCRADIO oRecord:AirBagDep VALUE 'Y'  PROMPT 'YES' COLOR 1,BD_SALMON ;
      PARENT oAirbagDep
		 @ 0,60 DCRADIO oRecord:AirBagDep VALUE 'N' PROMPT 'No' COLOR 1,BD_KEYLIME ;
      PARENT oAirbagdep

		 @ 2,1 DCGROUP oBodyWork CAPTION 'Has The Vehicle Had Body Or Paint Work  ' SIZE 80,1  PARENT oDisclose
		 @ 0,50 DCRADIO oRecord:BodyWork VALUE  'Y' PROMPT 'YES' COLOR 1,BD_SALMON ;
      PARENT oBodyWork
		 @ 0,60 DCRADIO oRecord:BodyWork VALUE 'N' PROMPT 'No' COLOR 1,BD_KEYLIME ;
      PARENT oBodywork

		 @ 3,1 DCGROUP oOdomacc CAPTION 'Is The Odometer Reading Accurate  ' SIZE 80,1  PARENT oDisclose
		 @ 0,50 DCRADIO oRecord:Odomacc VALUE 'Y' PROMPT 'YES' COLOR 1,BD_KEYLIME ;
      PARENT oOdomacc
		 @ 0,60 DCRADIO oRecord:Odomacc VALUE 'N' PROMPT 'No' COLOR 1,BD_SALMON ;
      PARENT oOdomacc

		 @ 4,1 DCGROUP oFramedam CAPTION 'Has The Vehicle Frame Ever been Damaged ' SIZE 80,1 PARENT oDisclose
		 @ 0,50 DCRADIO oRecord:Framedam VALUE 'Y' PROMPT 'YES' COLOR 1,BD_SALMON ;
      PARENT oFramedam
		 @ 0,60 DCRADIO oRecord:Framedam VALUE 'N' PROMPT 'No' COLOR 1,BD_KEYLIME ;
      PARENT oFramedam

		 @ 5,1 DCGROUP oRepreq CAPTION  'Are Repairs Required to Engine Components ' SIZE 80,1 PARENT oDisclose
		 @ 0,50 DCRADIO oRecord:Repreq VALUE 'Y' PROMPT 'YES' COLOR 1,BD_SALMON ;
      PARENT oRepreq
		 @ 0,60 DCRADIO oRecord:RepReq VALUE 'N' PROMPT 'No' COLOR 1,BD_KEYLIME ;
      PARENT oRepreq

		 @ 6,1 DCGROUP oAfterperf CAPTION 'Are There Any Aftermarket Performance Engine Parts ' SIZE 80,1 PARENT oDisclose
		 @ 0,50 DCRADIO oRecord:Afterperf VALUE 'Y' PROMPT 'YES' COLOR 1,BD_SALMON ;
      PARENT oAfterperf
		 @ 0,60 DCRADIO oRecord:Afterperf VALUE 'N' PROMPT 'No' COLOR 1,BD_KEYLIME ;
      PARENT oAfterperf


		IF MANAGER()
			@ 25,90 DCSAY 'Appraised Value       ' get oRecord:Acv picture '999999' PARENT oStatic
			@ 26,90 DCSAY 'Extra Lines           ' get oRecord:Lines picture '9'    PARENT oStatic
			@ 27,90 DCSAY 'Appraised By Initials ' get oRecord:Appraiser          PARENT oStatic
		ENDIF


		@ 29,85 DCPUSHBUTTON CAPTION 'Blank Form' ACTION {|| PRNTBLNKAPRSL()} SIZE 20,2


		@ 29,105 DCPUSHBUTTON CAPTION 'Print Screen' ACTION {|| DC_SCRN2CLIPBOARD(),PRINTCLIPBOARD()} SIZE 20,2

		@ 31,87 DCSAY 'Press OK Button to Print Appraisal  '

	DCREAD GUI MODAL TO xStatus ADDBUTTONS OPTIONS GETOPTIONS FIT TITLE 'UC Appraisal Info' EVAL{|o| SETAPPWINDOW(o)}
	IF ! xStatus
		RETURN NIL
	ENDIF

	// NOW UPDATE APPRAISAL RECORD



	IF lNewAppraisal
		UCAPPRSL->(DC_DBGATHER(oRecord,.t.))
		IF UCAPPRSL->(DC_RECLOCK())
			REPLACE UCAPPRSL->COLOR WITH cColor
			UCAPPRSL->(DBRUNLOCK())
		ENDIF

	ENDIF
	IF !lNewAppraisal
		UCAPPRSL->(DBGOTO(oRecord:Record_Number))
		UCAPPRSL->(DC_DBGATHER(oRecord))
		IF UCAPPRSL->(DC_RECLOCK())
			REPLACE UCAPPRSL->COLOR WITH cColor
			UCAPPRSL->(DBRUNLOCK())
		ENDIF
	ENDIF


	/// UPDATE NEWUPS RECORD OBJECT
	IF VALTYPE(oDeskRecord)='O'
	 oDeskRecord:TRADEYR:=oRecord:Year
	 oDeskRecord:TRADEDESC:=oRecord:Model
	 oDeskRecord:TRADEMAKE:=oRecord:Make
	 oDeskRecord:TRADEACV:=oRecord:ACV+(100*oRecord:Lines)
	 oDeskRecord:TRADEVIN:=oRecord:Vin
	 oDeskRecord:TRADECOLOR:=cColor
	 oDeskRecord:TRADEMILES:=oRecord:Mileage
	ENDIF


	SELECT NEWUPS
	NEWUPS->(DBGOTO(nRec))

	IF lFromBrowse
		NEWUPS->(DBGOTO(oDeskRecord:Record_Number))
		NEWUPS->(DC_DBGATHER(oDeskRecord))
	ENDIF

	IF !Print_Choice('Vehicle Appraisal', @nCopies,' ',.f.,@nOrient,'10.Courier New',@nDefMode,cOutfile)
	  RETURN NIL
   ENDIF
   IF !PrintOn('Vehicle Appraisal', oPrinter, nOrient, '10.Courier New', nCopies,.t.)
	  RETURN NIL
   ENDIF

	@ 0,0,66,132  dcprint BITMAP (GETFLG('ACCTPATH')+'UCAPPRAISAL.JPG')

	@ 3.5,113 DCPRINT SAY DATE() font '10.Courier New Bold'
	@ 9,20 DCPRINT SAY ALLTRIM(UCAPPRSL->LASTNAME) +','+ALLTRIM(UCAPPRSL->FIRSTNAME) font '10.Courier New Bold'
	@ 9,90 DCPRINT SAY NEWUPS->AREA+' '+NEWUPS->UPID   font '10.Courier New Bold'
	@ 10.5,22 DCPRINT SAY NEWUPS->STREET               font '10.Courier New Bold'
	@ 10.5,95 DCPRINT SAY UCAPPRSL->Model                       font '10.Courier New Bold'
	@ 12,22 DCPRINT SAY NEWUPS->CITYSTATE+' '+NEWUPS->ZIPCODE font '10.Courier New Bold'
	@ 12,90 DCPRINT SAY UCAPPRSL->Mileage picture '999999' font '10.Courier New Bold'
	@ 13.5,25 DCPRINT SAY UCAPPRSL->Year+'  '+UCAPPRSL->Make      font '10.Courier New Bold'
	@ 14,85 DCPRINT SAY TRANSFORM(UCAPPRSL->Vin,'@R X X X X X X X X X X X X X X X X X') font '10.Courier New Bold'
	@ 15,24 DCPRINT SAY cColor  font '10.Courier New Bold'
	@ 15,50 DCPRINT SAY UCAPPRSL->Interior  font '10.Courier New Bold'
	/// NOW PRINT OPTIONS


	@ 16,95 DCPRINT SAY 'Cylinders  '
	@ 16,120 DCPRINT SAY UCAPPRSL->CYLINDERS

	IF UCAPPRSL->OWNER='Y'
		@ 17,70 DCPRINT SAY 'YES!' font '12.Courier New Bold'
	else
		@ 17,70 DCPRINT SAY 'No'
	ENDIF

	@ 17,95 DCPRINT SAY 'EngineSize'
	@ 17,115 DCPRINT SAY UCAPPRSL->ENGSIZE

	IF UCAPPRSL->EXTCON='Y'
		@ 18.3,70 DCPRINT SAY 'YES!' font '12.Courier New Bold'
	else
		@ 18.3,70 DCPRINT SAY 'No'
	ENDIF

	@ 18,95 DCPRINT SAY 'TransMission '
	@ 18,115 DCPRINT SAY UCAPPRSL->TRANS
	@ 19.5,76 DCPRINT SAY  ALLTRIM(STR(UCAPPRSL->CURRPAY)) font '12.Courier New Bold'

	@ 19,95 DCPRINT SAY 'Speeds '
	@ 19,115 DCPRINT SAY UCAPPRSL->Speeds picture '9'
	IF UCAPPRSL->AWD
	 @ 20,95 DCPRINT SAY '*ALL WHEEL DRIVE ' font '10.Courier New Bold'
	 @ 20,120 DCPRINT SAY 'YES'               font '10.Courier New Bold'
	ELSE
	 @ 20,95 DCPRINT SAY 'All/4 Wheel Drive '
	 @ 20,120 DCPRINT SAY '___'
	ENDIF
	IF UCAPPRSL->INSTOCK='Y'
		@ 21,70 DCPRINT SAY 'YES!' font '12.Courier New Bold'
	else
		@ 21,70 DCPRINT SAY 'No'
	ENDIF

	@ 21,95 DCPRINT SAY 'Number of Doors'
	@ 21,120 DCPRINT SAY UCAPPRSL->NUMDOORS
	IF UCAPPRSL->WINDOWS
	 @ 22,95 DCPRINT SAY 'Power Windows '
	 @ 22,120 DCPRINT SAY 'YES'               font '10.Courier New Bold'
	ELSE
	 @ 22,95 DCPRINT SAY 'Power Windows '
	 @ 22,120 DCPRINT SAY '___'
	ENDIF
	IF UCAPPRSL->LOCKS
	 @ 23,95 DCPRINT SAY 'Power Door Locks '
	 @ 23,120 DCPRINT SAY 'YES'               font '10.Courier New Bold'
	ELSE
	 @ 23,95 DCPRINT SAY 'Power Door Locks '
	 @ 23,120 DCPRINT SAY '___'
	ENDIF
	IF UCAPPRSL->AC
	 @ 24,95 DCPRINT SAY 'Air Conditioning '
	 @ 24,120 DCPRINT SAY 'YES'               font '10.Courier New Bold'
	ELSE
	 @ 24,95 DCPRINT SAY 'Air Conditioning '
	 @ 24,120 DCPRINT SAY '___'
	ENDIF
	IF UCAPPRSL->ABS
	 @ 25,95 DCPRINT SAY 'Anti Lock Brakes '
	 @ 25,120 DCPRINT SAY 'YES'               font '10.Courier New Bold'
	ELSE
	 @ 25,95 DCPRINT SAY 'Anti Lock Brakes '
	 @ 25,120 DCPRINT SAY '___'
	ENDIF
	IF UCAPPRSL->CD
	 @ 26,95 DCPRINT SAY 'CD Player '
	 @ 26,120 DCPRINT SAY 'YES'               font '10.Courier New Bold'
	ELSE
	 @ 26,95 DCPRINT SAY 'CD Player '
	 @ 26,120 DCPRINT SAY '___'
	ENDIF

	IF UCAPPRSL->AIRBAGDEP='Y'
		@ 27,60 DCPRINT SAY 'AIR BAG DEPLOYED YES!' font '10.Courier New Bold'
	else
		@ 27,60 DCPRINT SAY 'Air Bag Deployed NO'
	ENDIF



	IF UCAPPRSL->DVD
	 @ 27,95 DCPRINT SAY '*DVD PLAYER*'       font '10.Courier New Bold'
	 @ 27,120 DCPRINT SAY 'YES'              font '10.Courier New Bold'
	ELSE
	 @ 27,95 DCPRINT SAY 'DVD Player '
	 @ 27,120 DCPRINT SAY '___'
	ENDIF

	IF UCAPPRSL->BODYWORK='Y'
		@ 28,60 DCPRINT SAY 'VEHICLE BODY WORK YES!' font '10.Courier New Bold'
	else
		@ 28,60 DCPRINT SAY 'Vehicle Body Work No'
	ENDIF




	IF UCAPPRSL->NAV
	 @ 28,95 DCPRINT SAY '*NAVIGATION*'       font '10.Courier New Bold'
	 @ 28,120 DCPRINT SAY 'YES'              font '10.Courier New Bold'
	ELSE
	 @ 28,95 DCPRINT SAY 'Navigation '
	 @ 28,120 DCPRINT SAY '___'
	ENDIF

	IF UCAPPRSL->ODOMACC='Y'
		@ 29,60 DCPRINT SAY 'Odometer Accurate Yes'
	else
		@ 29,60 DCPRINT SAY 'ODOMETER ACCURATE NO!'  font '10.Courier New Bold'
	ENDIF




	IF UCAPPRSL->HEATSEAT
	 @ 29,95 DCPRINT SAY '*HEATED SEATS*'       font '10.Courier New Bold'
	 @ 29,120 DCPRINT SAY 'YES'              font '10.Courier New Bold'
	ELSE
	 @ 29,95 DCPRINT SAY 'Heated Seats '
	 @ 29,120 DCPRINT SAY '___'
	ENDIF

	IF UCAPPRSL->FRAMEDAM='Y'
		@ 30,60 DCPRINT SAY 'FRAME DAMAGE YES!' font '10.Courier New Bold'
	else
		@ 30,60 DCPRINT SAY 'Frame Damage No'
	ENDIF



	IF UCAPPRSL->HYBRID
	 @ 30,95 DCPRINT SAY '*HYBRID*'       font '10.Courier New Bold'
	 @ 30,120 DCPRINT SAY 'YES'              font '10.Courier New Bold'
	ELSE
	 @ 30,95 DCPRINT SAY 'Hybrid '
	 @ 30,120 DCPRINT SAY '___'
	ENDIF

	IF UCAPPRSL->REPREQ='Y'
		@ 31,60 DCPRINT SAY 'REPAIRS REQUIRED YES!' font '10.Courier New Bold'
	else
		@ 31,60 DCPRINT SAY 'Repairs Required No'
	ENDIF



	IF UCAPPRSL->SUNROOF
	 @ 31,95 DCPRINT SAY '*SUN ROOF*'       font '10.Courier New Bold'
	 @ 31,120 DCPRINT SAY 'YES'              font '10.Courier New Bold'
	ELSE
	 @ 31,95 DCPRINT SAY 'Sun Roof '
	 @ 31,120 DCPRINT SAY '___'
	ENDIF

	IF UCAPPRSL->AFTERPERF='Y'
		@ 32,60 DCPRINT SAY 'AFTERMARKET PARTS YES!' font '10.Courier New Bold'
	else
		@ 32,60 DCPRINT SAY 'After Market Parts No'
	ENDIF



	IF UCAPPRSL->REARCAM
	 @ 32,95 DCPRINT SAY '*REAR CAMERA*'       font '10.Courier New Bold'
	 @ 32,120 DCPRINT SAY 'YES'              font '10.Courier New Bold'
	ELSE
	 @ 32,95 DCPRINT SAY 'Rear Camera '
	 @ 32,120 DCPRINT SAY '___'
	ENDIF



	IF UCAPPRSL->ONSTAR
	 @ 33,95 DCPRINT SAY '*ONSTAR*'       font '10.Courier New Bold'
	 @ 33,120 DCPRINT SAY 'YES'              font '10.Courier New Bold'
	ELSE
	 @ 33,95 DCPRINT SAY 'On Star '
	 @ 33,120 DCPRINT SAY '___'
	ENDIF


	IF UCAPPRSL->SIDEBAGS
	 @ 34,95 DCPRINT SAY '*SIDE AIR BAGS*'       font '10.Courier New Bold'
	 @ 34,120 DCPRINT SAY 'YES'              font '10.Courier New Bold'
	ELSE
	 @ 34,95 DCPRINT SAY 'Side Air Bags '
	 @ 34,120 DCPRINT SAY '___'
	ENDIF



	IF UCAPPRSL->LEATHER
	 @ 35,95 DCPRINT SAY '*LEATHER INTERIOR*'       font '10.Courier New Bold'
	 @ 35,120 DCPRINT SAY 'YES'              font '10.Courier New Bold'
	ELSE
	 @ 35,95 DCPRINT SAY 'Leather Interior '
	 @ 35,120 DCPRINT SAY '___'
	ENDIF



	IF UCAPPRSL->SATRAD
	 @ 36,95 DCPRINT SAY '*SATELLITE RADIO*'       font '10.Courier New Bold'
	 @ 36,120 DCPRINT SAY 'YES'              font '10.Courier New Bold'
	ELSE
	 @ 36,95 DCPRINT SAY 'Satellite Radio '
	 @ 36,120 DCPRINT SAY '___'
	ENDIF


	IF UCAPPRSL->MP3
	 @ 37,95 DCPRINT SAY '*MP3 PLAYER*'       font '10.Courier New Bold'
	 @ 37,120 DCPRINT SAY 'YES'              font '10.Courier New Bold'
	ELSE
	 @ 37,95 DCPRINT SAY 'MP3 Player '
	 @ 37,120 DCPRINT SAY '___'
	ENDIF
	IF UCAPPRSL->BIGWHEELS
	 @ 38,95 DCPRINT SAY '*20 INCH WHEELS*'       font '10.Courier New Bold'
	 @ 38,120 DCPRINT SAY 'YES'              font '10.Courier New Bold'
	ELSE
	 @ 38,95 DCPRINT SAY '20 Inch Wheels '
	 @ 38,120 DCPRINT SAY '___'
	ENDIF

	IF !empty(UCAPPRSL->OTHER1)
	 @ 34,60 DCPRINT SAY 'Also Has '+ UCAPPRSL->OTHER1
	ENDIF
	IF !empty(UCAPPRSL->OTHER2)
	 @ 35,60 DCPRINT SAY  'Also Has '+ UCAPPRSL->OTHER2
	ENDIF
	IF !empty(UCAPPRSL->OTHER3)
	 @ 36,60 DCPRINT SAY  'Also Has '+ UCAPPRSL->OTHER3
	ENDIF
	IF !empty(UCAPPRSL->OTHER4)
	 @ 37,60 DCPRINT SAY  'Also Has '+ UCAPPRSL->OTHER4
	ENDIF
	IF !empty(UCAPPRSL->OTHER5)
	 @ 38,60 DCPRINT SAY  'Also Has '+ UCAPPRSL->OTHER5
	ENDIF
	@ 40,30 DCPRINT SAY UCAPPRSL->Comments
	IF UCAPPRSL->ACV > 0
	 @ 59,90 DCPRINT SAY 'Appraised Value '
	 @ 59,110 DCPRINT SAY UCAPPRSL->ACV PICTURE '99999'
	 IF UCAPPRSL->LINES > 0
		FOR I := 1 TO UCAPPRSL->LINES
			@ DC_PRINTERROW()+1,110 DCPRINT SAY '_________'
			nAcv:=nAcv+100
		NEXT
	 ENDIF
	ENDIF



	printoff(oPrinter)

	/*
	select NEWUPS
	IF NEWUPS->(DC_RECLOCK())
		REPLACE NEWUPS->TRADEDESC WITH cModel
		REPLACE NEWUPS->TRADEYR WITH cYear
		REPLACE NEWUPS->TRADEMAKE WITH cMake
		REPLACE NEWUPS->TRADECOLOR WITH cColor
		REPLACE NEWUPS->TRADEMILES WITH nMiles
		REPLACE NEWUPS->TRADEVIN WITH cVin
		REPLACE NEWUPS->TRADEACV WITH nAcv
		NEWUPS->(DBRUNLOCK(nRec))
	ENDIF                            */
	RETURN NIL



	FUNCTION prntblnkaprsl()

		local nCopies:=1,nOrient:=1,nDefmode:=4, oPrinter,cOutfile:=''
		local getlist:={}
		IF !Print_Choice('Vehicle Appraisal', @nCopies,' ',.f.,@nOrient,'10.Courier New',@nDefMode,cOutfile)
   	  RETURN NIL
      ENDIF
      IF !PrintOn('Vehicle Appraisal', oPrinter, nOrient, '10.Courier New', nCopies,.t.)
	     RETURN NIL
      ENDIF


		@ 0,0,66,132  dcprint BITMAP (GETFLG('ACCTPATH')+'UCAPPRAISAL.JPG')

	@ 16,95 DCPRINT SAY 'Cylinders  '
	@ 16,120 DCPRINT SAY '___'
	@ 17,70 DCPRINT SAY '___'

	@ 17,95 DCPRINT SAY 'EngineSize'
	@ 17,120 DCPRINT SAY '___'

	@ 18,70 DCPRINT SAY '___'

	@ 18,95 DCPRINT SAY 'TransMission '
	@ 18,120 DCPRINT SAY '___'

	@ 19,95 DCPRINT SAY 'Speeds '
	@ 19,120 DCPRINT SAY '___'

	@ 20,70 DCPRINT SAY '___'

	@ 20,95 DCPRINT SAY '*ALL WHEEL DRIVE ' font '10.Courier New Bold'
	@ 20,120 DCPRINT SAY '___'

	@ 21,95 DCPRINT SAY 'Number of Doors'
	@ 21,120 DCPRINT SAY '___'

	@ 22,95 DCPRINT SAY 'Power Windows '
	@ 22,120 DCPRINT SAY '___'

	@ 23,95 DCPRINT SAY 'Power Door Locks '
	@ 236,120 DCPRINT SAY '___'

	@ 24,95 DCPRINT SAY 'Air Conditioning '
	@ 24,120 DCPRINT SAY '___'

	@ 25,95 DCPRINT SAY 'Anti Lock Brakes '
	@ 25,120 DCPRINT SAY '___'

	@ 26,95 DCPRINT SAY 'CD Player '
	@ 26,120 DCPRINT SAY '___'
	@ 27,50 DCPRINT SAY 'AIR BAG DEPLOYED  Yes / No' font '10.Courier New Bold'

	@ 27,95 DCPRINT SAY '*DVD PLAYER*'       font '10.Courier New Bold'
	@ 27,120 DCPRINT SAY '___'


	@ 28,50 DCPRINT SAY 'VEHICLE BODY WORK Yes / No' font '10.Courier New Bold'

	@ 28,95 DCPRINT SAY '*NAVIGATION*'       font '10.Courier New Bold'
	@ 28,120 DCPRINT SAY '___'

	@ 29,50 DCPRINT SAY 'ODOMETER ACCURATE Yes / No'  font '10.Courier New Bold'

	@ 29,95 DCPRINT SAY '*HEATED SEATS*'       font '10.Courier New Bold'
	@ 29,120 DCPRINT SAY '___'

	@ 30,50 DCPRINT SAY 'FRAME DAMAGE      Yes / No' font '10.Courier New Bold'

	@ 30,95 DCPRINT SAY '*HYBRID*'       font '10.Courier New Bold'
	@ 30,120 DCPRINT SAY '___'


	@ 31,50 DCPRINT SAY 'REPAIRS REQUIRED  Yes / No' font '10.Courier New Bold'

	@ 31,95 DCPRINT SAY '*SUN ROOF*'       font '10.Courier New Bold'
	@ 31,120 DCPRINT SAY '___'


	@ 32,50 DCPRINT SAY 'AFTERMARKET PARTS Yes / No' font '10.Courier New Bold'
	@ 32,95 DCPRINT SAY '*REAR CAMERA*'       font '10.Courier New Bold'
	@ 32,120 DCPRINT SAY '___'

	@ 33,95 DCPRINT SAY '*ONSTAR*'       font '10.Courier New Bold'
	@ 33,120 DCPRINT SAY '___'


	@ 34,95 DCPRINT SAY '*SIDE AIR BAGS*'       font '10.Courier New Bold'
	@ 34,120 DCPRINT SAY '___'              font '10.Courier New Bold'



   @ 35,95 DCPRINT SAY '*LEATHER INTERIOR*'       font '10.Courier New Bold'
   @ 35,120 DCPRINT SAY '___'              font '10.Courier New Bold'



   @ 36,95 DCPRINT SAY '*SATELLITE RADIO*'       font '10.Courier New Bold'
   @ 36,120 DCPRINT SAY '___'              font '10.Courier New Bold'

   @ 37,95 DCPRINT SAY '*MP3 PLAYER*'       font '10.Courier New Bold'
   @ 37,120 DCPRINT SAY '___'              font '10.Courier New Bold'
   @ 38,95 DCPRINT SAY '*20 INCH WHEELS*'       font '10.Courier New Bold'
   @ 38,120 DCPRINT SAY '___'              font '10.Courier New Bold'

	@ 34,50 DCPRINT SAY 'Also Has_______________________ '
	@ 35,50 DCPRINT SAY 'Also Has_______________________ '
	@ 36,50 DCPRINT SAY 'Also Has_______________________ '
	@ 37,50 DCPRINT SAY 'Also Has_______________________ '
	@ 38,50 DCPRINT SAY 'Also Has_______________________ '



	PRINTOFF(oPrinter)

	RETURN NIL


	STATIC FUNCTION _UPDATEAPPRSL(cVin,cMake,cModel,cYear,cColor,nMiles,cInterior,nAcv,cManagerInit,nLines,dDate,cEngsize,cCyl,lLocks,lWindows,cTrans,;
	  nSpeeds,lAc,lAwd,lAbs,lCd,lSidebags,lDVD,lNav,lHeatseat,lOnStar,lRearcam,lLeather,lMP3,lSatrad,lBigWheels,lSunroof,lPowseat,;
	  cOther,cOther2,cOther3,cOther4,cOther5,cLastName,cFirstname,cSeries,cComments,clAirbagdep,clBodywork,clOdomacc,clFramedam,;
	  clRepreq,clAfterperf,cSM,lHybrid, nNumdoors,clOwner,clExtcon,clInstock,nCurrpay)
	  local nrec:=0
	IF UCAPPRSL->(dc_reclock())
		nREC:=UCAPPRSL->(RECNO())
		REPLACE UCAPPRSL->VIN WITH cVin
		REPLACE UCAPPRSL->MAKE WITH cMake
		REPLACE UCAPPRSL->MODEL with cModel
		REPLACE UCAPPRSL->YEAR with cYear
		REPLACE UCAPPRSL->COLOR  with cColor
		REPLACE UCAPPRSL->MILEAGE  with nMiles
		REPLACE UCAPPRSL->INTERIOR  with cInterior
		REPLACE UCAPPRSL->ACV  with nAcv
		REPLACE UCAPPRSL->APPRAISER  with cManagerInit
		REPLACE UCAPPRSL->LINES   with nLines
		REPLACE UCAPPRSL->APPRDATE with dDate
		REPLACE UCAPPRSL->ENGSIZE  with cEngsize
		REPLACE UCAPPRSL->CYLINDERS   with cCyl
		REPLACE UCAPPRSL->LOCKS     with lLocks
		REPLACE UCAPPRSL->WINDOWS with lWindows
		REPLACE UCAPPRSL->TRANS     with cTrans
		REPLACE UCAPPRSL->SPEEDS   with nSpeeds
		REPLACE UCAPPRSL->AC           with lAc
		REPLACE UCAPPRSL->AWD         with lAwd
		REPLACE UCAPPRSL->ABS         with lAbs
		REPLACE UCAPPRSL->CD           with lCd
		REPLACE UCAPPRSL->SIDEBAGS  with lSidebags
		REPLACE UCAPPRSL->DVD          with lDVD
		REPLACE UCAPPRSL->NAV          with lNav
		REPLACE UCAPPRSL->HEATSEAT  with lHeatseat
		REPLACE UCAPPRSL->ONSTAR      with lOnstar
		REPLACE UCAPPRSL->REARCAM    with lRearcam
		REPLACE UCAPPRSL->LEATHER    with lLeather
		REPLACE UCAPPRSL->MP3            with lMP3
		REPLACE UCAPPRSL->SATRAD      with lSatrad
		REPLACE UCAPPRSL->BIGWHEELS  with lBigwheels
		REPLACE UCAPPRSL->SUNROOF      with lSunroof
		REPLACE UCAPPRSL->POWSEAT      with lPowseat
		REPLACE UCAPPRSL->OTHER          with cOther
		REPLACE UCAPPRSL->OTHER2        with cOther2
		REPLACE UCAPPRSL->OTHER3        with cOther3
		REPLACE UCAPPRSL->OTHER4        with cOther4
		REPLACE UCAPPRSL->OTHER5        with cOther5
		REPLACE UCAPPRSL->LASTNAME    with cLastName
		REPLACE UCAPPRSL->FIRSTNAME  with cFirstname
		REPLACE UCAPPRSL->SERIES        with cSeries
		REPLACE UCAPPRSL->COMMENTS    with cComments
		REPLACE UCAPPRSL->AIRBAGDEP  with clAirbagdep
		REPLACE UCAPPRSL->BODYWORK    with clBodywork
		REPLACE UCAPPRSL->ODOMACC      with clodomacc
		REPLACE UCAPPRSL->FRAMEDAM    with clFramedam
		REPLACE UCAPPRSL->REPREQ        with clRepreq
		REPLACE UCAPPRSL->AFTERPERF  with clAfterperf
		REPLACE UCAPPRSL->SM  with cSM
		REPLACE UCAPPRSL->HYBRID WITH lHybrid
		REPLACE UCAPPRSL->NUMDOORS with  nNumdoors
		REPLACE UCAPPRSL->OWNER WITH clOwner
		REPLACE UCAPPRSL->Extcon with clExtcon
		REPLACE UCAPPRSL->Instock with clInstock
		REPLACE UCAPPRSL->CURRPAY with nCurrpay
	 ENDIF
	 return nil


STATIC FUNCTION UpdateUpRec(aUpChars,aUpNums,aUpdates,aUpBoolean,nDowndeal,dDatesold,cMiddleinit,cCoDeath,cUpstat,clAfterdel,cBillStat)
	local nRec:=NEWUPS->(Recno())
	DEFAULT clAfterdel:=''
	if NEWUPS->(dc_reclock())
		REPLACE NEWUPS->UPID WITH aUpChars[UPSC_CUPID]
		REPLACE NEWUPS->AREA WITH aUpChars[UPSC_CAREACODE]
		REPLACE NEWUPS->FIRSTNAME WITH aUpChars[UPSC_CFIRSTNAME]
		REPLACE NEWUPS->LASTNAME WITH aUpChars[UPSC_CLASTNAME]
		REPLACE NEWUPS->STREET WITH aUpChars[UPSC_CSTREET]
		REPLACE NEWUPS->CITYSTATE WITH aUpChars[UPSC_CCITYSTATE]
		REPLACE NEWUPS->ZIPCODE WITH aUpChars[UPSC_CZIPCODE]
		REPLACE NEWUPS->INTEREST WITH aUpChars[UPSC_CVEHINTEREST]
		REPLACE NEWUPS->UPTYPE WITH aUpNums[UPSN_NUPTYPE]
		REPLACE NEWUPS->NEWUSED WITH aUpChars[UPSC_CNEWUSED]
		REPLACE NEWUPS->SALESMAN WITH aUpChars[UPSC_CSALESPERSON]
		REPLACE NEWUPS->COMMENTS WITH aUpChars[UPSC_CCOMMENT]
		REPLACE NEWUPS->COMMENT2 WITH aUpChars[UPSC_CCOMMENT2]
		REPLACE NEWUPS->SALECODE WITH aUpNums[UPSN_NSOLDCODE]
		REPLACE NEWUPS->TOMGR WITH aUpChars[UPSC_CTOMGR]
		REPLACE NEWUPS->STOCKNUM WITH aUpChars[UPSC_CSTOCKNO]
		REPLACE NEWUPS->ADSOURCE WITH aUpChars[UPSC_CADSOURCE]
		REPLACE NEWUPS->SALEDATE WITH dDateSold
		REPLACE NEWUPS->DEMO WITH aUpChars[UPSC_CLDEMOYN]
		REPLACE NEWUPS->EMAIL WITH aUpChars[UPSC_CEMAIL]
		REPLACE NEWUPS->DEMOSTK WITH aUpChars[UPSC_CDEMODSTKNO]
		REPLACE NEWUPS->WORKPHN WITH aUpChars[UPSC_CWORKPHONE]
		REPLACE NEWUPS->ORIGSOURCE WITH aUpChars[UPSC_ORIGSOURCE]
		REPLACE NEWUPS->TICKLE WITH aUpDates[UPSD_DTICKLEDATE]
		REPLACE NEWUPS->FRANCHISE WITH aUpChars[UPSC_CFRANCHISE]
		REPLACE NEWUPS->CELLPHONE WITH aUpChars[UPSC_CCELLPHONE]
		REPLACE NEWUPS->UPSTATUS WITH cUpstat
		REPLACE NEWUPS->APPOINT WITH aUpBoolean[UPSL_APPOINT]
		REPLACE NEWUPS->WRITEUP WITH aUpBoolean[UPSL_WRITEUP]
		REPLACE NEWUPS->TOED WITH aUpBoolean[UPSL_TOED]
		REPLACE NEWUPS->DEMOED WITH aUpBoolean[UPSL_DEMOED]
		REPLACE NEWUPS->DATEIN WITH aUpDates[UPSD_DDATEIN]
		REPLACE NEWUPS->PREFMETHOD WITH aUpChars[UPSC_PREFMETH]

		IF NEWUPS->DEMO='Y'
			REPLACE NEWUPS->DEMOED WITH TRUE
		ENDIF

		REPLACE NEWUPS->BEBACK3 WITH aUpDates[UPSD_DORIGDTIN]  /// BEBACK 3 IS NOW THE ORIGINAL DATE IN
		IF clAfterdel='Y'
		 REPLACE NEWUPS->SENDLTR WITH 'Y'
		ENDIF
		REPLACE NEWUPS->BILLSTAT WITH cBillStat
						/// IF THIS IS A SOLD VEHICLE
		IF cUpstat='S'
			REPLACE NEWUPS->SOLD WITH TRUE
			IF !empty(aUpDates[UPSD_DBEBACK1])
				REPLACE NEWUPS->SALECODE WITH 4
			ELSE
				REPLACE NEWUPS->SALECODE WITH aUpNums[UPSN_NUPTYPE]
			ENDIF
		ENDIF
		IF !empty(aUpDates[UPSD_DBEBACK1])
				replace NEWUPS->DATEIN WITH aUpDates[UPSD_DBEBACK1]
		ENDIF
		IF !empty(aUpDates[UPSD_DBEBACK2])
				replace NEWUPS->DATEIN WITH aUpDates[UPSD_DBEBACK2]
		ENDIF

		REPLACE NEWUPS->BEBACK1 WITH aUpDates[UPSD_DBEBACK1]
		REPLACE NEWUPS->BEBACK2 WITH aUpDates[UPSD_DBEBACK2]

		REPLACE NEWUPS->downdeal with nDowndeal
		replace NEWUPS->saledate with dDatesold
		replace NEWUPS->causedead with cCoDeath
		replace NEWUPS->Middleinit with cMiddleinit
		NEWUPS->(dbrunlock(nRec))
	endif
 return nil

/// FUNCTION TO INSURE ONLY GOOD DATA GETS IN
Static Function editupstat(oRecord)
	if !oRecord:Upstatus $('ASRFD')
		RETURN FALSE
	endif
 	IF oRecord:Upstatus='S'
	   	oRecord:SOLD:=.T.
  	ELSE
		oRecord:SOLD:=.F.
 	ENDIF

	if oRecord:Upstatus ='D'
		IF !MANAGER()
			DC_WINALERT('You do not have the authority over the life and death of an UP!')
		   return .f.
	   endif
   endif
return .T.


static FUNCTION BuyersOrder(nRec,lFinal,oRecord)
	local getlist:={}, getoptions, xstatus , lPrevUE , lPdf , cMemo, cMemoline , nLinecount , I , lPrintTerms , cCell, cWork , lSkip:=.t. , lEspanol:=.f.
	LOCAL cMake,cModel,cSeries,cVehicleID, cColor, cState, cCity , cYear, nMileage, cInterior, cSmName,cSmName2:='',cType,cFmrLR:='N'
	LOCAL cOutfile:='c:\BuyersOrder.pdf' , lPrintExtra:=.t. ,lPrintSupplement:=.t., nNumCopy:=1 ,lPrintList:=.t.,lPrintdate:=.t.,oPrinter  , lLocks:=.t.,lMats:=.t.,lHasLocks:=.f. ,lHasMats:=.f.
	default lFinal:=.f.
	DCGEToptions saywidth 0 sayfont '12.Courier New' getfont '12.Courier New' autoresize
	/// establish default values
	cMake:=cModel:=cSeries:=cVehicleID:=cColor:= cState:= cCity:= cYear:= cInterior:=cTYPE:=cSmName:=' '

	nMileage:=0
	cVehicleId:=NEWUPS->Quotevin
	cMake:=NEWUPS->VEHQUOTED
	cYear:=NEWUPS->YRQUOTED

	 //IF lFinal
	  //	nanna()
	  //	return nil
		//DC_WINALERT('I will be shutting down the ability to print a final buyers order from UPS on Monday;March 31. You will only be able to print from the credit app system.; If you want to continue to hear the last tones just keep on using the old verion. ')
	 //ENDIF

	// get vehicle info
	IF !empty(NEWUPS->STKQUOTED)
		IF NEWUPS->NEWUSED='N'
			GETNEWINFO(NEWUPS->STKQUOTED,@cMake,@cModel,@cSeries,@cVehicleId,@cYear,@cColor,@nMileage,@cInterior,@cType)
		ELSE
			GETUSEDINFO(NEWUPS->STKQUOTED,@cMake,@cModel,@cSeries,@cVehicleId,@cYear,@cColor,@nMileage,@cInterior,@cType,@cFmrLR)
		ENDIF
	ENDIF

	IF EMPTY(cMake)
		cMake:=NEWUPS->VEHQUOTED
		cYear:=NEWUPS->YRQUOTED
	ENDIF

	//get separate city and state
	getzip(NEWUPS->ZipCode,@cCity,@cState)
	/// get salesperson name
	getsmname(NEWUPS->SALESMAN,@cSmName)
	IF !empty(NEWUPS->SM2)
	  getsmname(NEWUPS->SM2,@cSmName2)
	ENDIF

	select NEWUPS
	NEWUPS->(DBGOTO(nRec))

	lPrevue:=.f.
	lPdf:=.f.
	lPrintTerms:=.F.
	lPrintExtra:=.F.
	lPrintlist:=.f.
	lEspanol:=.f.


	@ 5,5 DCSAY 'Additional Printer Choices'
	@ 6,5 DCCHECKBOX lPrintDate PROMPT 'Un Check Here to not Print Date/Time stamp on Buyers Order'
	@ 8,5 DCCHECKBOX lPdf PROMPT 'Check Here to Print to PDF File for eMail'
	@ 10,5 DCCHECKBOX lPrevue PROMPT 'Check Here to Preview on Screen'
	@ 12,5 DCCHECKBOX lPrintTerms PROMPT 'Check Here to Print Page 2 Terms'
	@ 14,5 DCCHECKBOX lPrintExtra PROMPT  'Check Here to Print Extra Copy for Customer'
	@ 16,5 DCCHECKBOX lPrintList PROMPT 'Check Here to have List Price Print on Buyers Order'
	@ 17,5 DCCHECKBOX lPrintSupplement  PROMPT 'Check Here to have Delivery Check List Print'
	@ 18,5 DCCHECKBOX lEspanol  PROMPT 'Haga clic aquí para imprimir documentos en español'


	dcread gui modal to xstatus addbuttons fit options getoptions title 'Printer Choices' eval{|o| setappwindow(o)}
	IF !xStatus
		RETURN NIL
	ENDIF

	IF lPrintExtra
		nNumCopy:=2
	ENDIF

	if lPDF
		/// check to see if pdf driver is installed
		if AScan(XbpPrinter():new():list(),{|c|Upper(c)=='WIN2PDF'})=0
			dc_winalert('You do not have the WIN2PDF Driver. You must download it from meadowlandsystems.com')
			return nil
		endif
		DCPRINT ON;
			NAME 'Win2PDF';
			TOFILE ;
			OUTFILE (cOutfile) OVERWRITE ;
			SIZE 66,132 ;
			FONT '10.Courier New'
	ENDIF

	IF !lPDF
	 DCPRINT ON SIZE 66,132 TO oPrinter FONT '12.Courier New Bold' _PREVIEW lPrevue ZOOMFACTOR 1.4 PSIZE 1024,768
	 IF VALTYPE(OPRINTER) # 'O' .OR. !OPRINTER:LACTIVE
				RETURN NIL
	 ENDIF
	ENDIF

	DO WHILE nNUMCOPY > 0

		IF lEspanol
			@ 0,0,66,133 dcprint bitmap (GETFLG('ACCTPATH')+'buyersordspanish.jpg')
		 else
		   @ 0,0,66,133 dcprint bitmap (GETFLG('ACCTPATH')+'buyersord.jpg')
		ENDIF
		IF lPrintDate
		 @ 1,102 DCPRINT SAY DTOC(DATE())+' '+TIME()
		else
		 @ 1,102 DCPRINT SAY ''
		endif
		/// IF THIS IS A SPLIT PRINT BOTH SM NAMES
		IF !EMPTY(NEWUPS->SM2)
		 @ DC_PRINTERROW()+4.1,89 DCPRINT SAY 'COSALE!!'
		 @ DC_PRINTERROW(),102 DCPRINT SAY NEWUPS->SM2+' '+cSmName2
		 @ DC_PRINTERROW()+1,14 DCPRINT SAY ALLTRIM(NEWUPS->LASTNAME)+','+ALLTRIM(NEWUPS->FIRSTNAME)+'  '+NEWUPS->MiddleInit
		 @ DC_PRINTERROW(),102  DCPRINT SAY NEWUPS->SALESMAN+' '+cSmName
		ELSE           // IF NO SPLIT THEN DO NORMAL PRINT
		 @ DC_PRINTERROW()+5,14 DCPRINT SAY ALLTRIM(NEWUPS->LASTNAME)+','+ALLTRIM(NEWUPS->FIRSTNAME)+'  '+NEWUPS->MiddleInit
		 @ DC_PRINTERROW(),102  DCPRINT SAY NEWUPS->SALESMAN+' '+cSmName
		ENDIF
		@ DC_PRINTERROW()+1,100  DCPRINT SAY NEWUPS->AREA+'-'+NEWUPS->UPID
		@ DC_PRINTERROW()+.5,14   DCPRINT SAY ALLTRIM(NEWUPS->STREET)
		cCell:=NEWUPS->CELLPHONE
		cWork:=NEWUPS->workphn
		_formatphn(@cCell)
		_formatphn(@cWork)
		@ dc_printerrow()+.5,97 DCPRINT SAY 'W '+cWork+' '+'C '+cCell font '8.Courier New Bold'
		@ DC_PRINTERROW()+1,14   DCPRINT SAY ALLTRIM(cCity)
		@ DC_PRINTERROW(),60  DCPRINT SAY cState
   	@ DC_PRINTERROW(),75  DCPRINT SAY NEWUPS->ZIPCODE
		@ DC_PRINTERROW(),93  DCPRINT SAY NEWUPS->EMAIL font '10.Courier New Bold'
		@ DC_PRINTERROW()+1.1,5 DCPRINT SAY 'Register To '+ NEWUPS->EXTRACHAR2    /// REGISTRATION INFO
		@ DC_PRINTERROW(),90 DCPRINT SAY 'Customer # ' + NEWUPS->REBDESC6   /// this field in the dbf is now used for customer number

		@ DC_PRINTERROW()+4.2,7  DCPRINT SAY cYear
		IF NEWUPS->NEWUSED='N'
			@ DC_PRINTERROW(),18 DCPRINT SAY 'NEW'
		ELSE
			@ DC_PRINTERROW(),18 DCPRINT SAY 'USED'
		ENDIF
		@ DC_PRINTERROW(),60 DCPRINT SAY cMake
		@ DC_PRINTERROW(),85 DCPRINT SAY cModel
		@ DC_PRINTERROW(),110 DCPRINT SAY cSeries
		@ DC_PRINTERROW()+1.2,7 DCPRINT SAY cType
		@ DC_PRINTERROW(),28 DCPRINT SAY cColor FONT '10.Courier New Bold'
		@ DC_PRINTERROW(),60 DCPRINT SAY cInterior FONT '10.Courier New Bold'
		IF len(alltrim(cVehicleId))=17
		  @ DC_PRINTERROW(),77 DCPRINT SAY TRANSFORM(cVehicleID,'@R X X X X X X X X X X X X X X X X X')
		endif
		@ DC_PRINTERROW()+5.3,5 DCPRINT SAY NEWUPS->Delivdate
		@ DC_PRINTERROW(),20 DCPRINT SAY NEWUPS->Deltime
		@ DC_PRINTERROW(),70 DCPRINT SAY NEWUPS->QUOTEMILES picture '999999'
		@ DC_PRINTERROW(),100 DCPRINT SAY NEWUPS->StkQuoted
		IF cFmrLR='Y'
			@ DC_PRINTERROW()+2.4,18 DCPRINT SAY  'XX'
		else
			@ DC_PRINTERROW()+2.4,20 DCPRINT SAY  ' '
		ENDIF

		lSkip:=.t.

	  IF lSkip                                       /// move line 1 down
				@ DC_PRINTERROW()+1,20 DCPRINT SAY ' '
				lSkip:=.f.
	  ENDIF


		IF NEWUPS->QUOTETYPE='C'
		 @ DC_PRINTERROW(),100 DCPRINT SAY  'CASH DEAL'
		ENDIF
		IF NEWUPS->QUOTETYPE='R'
		 @ DC_PRINTERROW(),100 DCPRINT SAY  'FINANCE DEAL'
		ENDIF
		IF NEWUPS->QUOTETYPE='L'
		 @ DC_PRINTERROW(),5 DCPRINT SAY 'Notice: Quoted Lease Payments Are Subject To Top Credit Rating Approval'  FONT '10.Arial Bold'
		 @ DC_PRINTERROW(),100 DCPRINT SAY  'LEASE DEAL'
		ENDIF
		IF lprintList
		 @ DC_PRINTERROW()+1,20 DCPRINT SAY 'List Price$'
		 @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->QUOTEMSRP PICTURE '999,999.99'
		ENDIF


		//PRICING SECTION
		IF NEWUPS->QUOTETYPE $('RC')
		 IF lEspanol
			@ DC_PRINTERROW()+1.2,20 DCPRINT SAY  'Su Precio $'
			else
		   @ DC_PRINTERROW()+1.2,20 DCPRINT SAY  'Your Price $'
		 endif
		 @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->QUOTED PICTURE '999,999.99'
		 IF lFinal
		  @ DC_PRINTERROW(),111 DCPRINT SAY NEWUPS->QUOTED PICTURE '999,999.99'
		 ENDIF
		ENDIF  //// quotetype RC
		/*
		IF ALLTRIM(NEWUPS->EXTRACHAR3) ='Y'
		 @ DC_PRINTERROW()+1.2,10 DCPRINT SAY  'Property Tax Included'
		 @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->EXTRANUM3 PICTURE '999,999.99'
		ENDIF    //// quotetype    */
		// SKIP LINE
		@ DC_PRINTERROW()+1.1,0 DCPRINT SAY ''
		IF NEWUPS->TRADEQUOTE > 0
			IF lEspanol
				@ DC_PRINTERROW(),20 DCPRINT SAY  'Subsidio de Comercio'
			  else
		      @ DC_PRINTERROW(),20 DCPRINT SAY  'Trade Allowance'
			endif
		 @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->TradeQuote PICTURE '999,999.99'
		ENDIF
		IF NEWUPS->QUOTETYPE $('RC')
			IF NEWUPS->REBATE1 > 0
			  @ DC_PRINTERROW(),68 DCPRINT SAY NEWUPS->REBDESC1 FONT '6.Courier New Bold'
			  @ DC_PRINTERROW(),79 DCPRINT SAY NEWUPS->REBATE1 PICTURE '9,999.99'  FONT '6.Courier New Bold'
			  @ DC_PRINTERROW(),87 DCPRINT SAY NEWUPS->REBCODE1  FONT '6.Courier New Bold'
			ENDIF
		else
			IF NEWUPS->LREBATE1 > 0
			  @ DC_PRINTERROW(),68 DCPRINT SAY NEWUPS->LREBDESC1 FONT '6.Courier New Bold'
			  @ DC_PRINTERROW(),79 DCPRINT SAY NEWUPS->LREBATE1 PICTURE '9,999.99'  FONT '6.Courier New Bold'
			  @ DC_PRINTERROW(),87 DCPRINT SAY NEWUPS->LREBCODE1  FONT '6.Courier New Bold'
			ENDIF
		ENDIF

		IF lFinal
			 IF ABS(NEWUPS->aftsaleam1) > 0
			  @ DC_PRINTERROW(),93 DCPRINT SAY NEWUPS->aftsale1 FONT '8.Courier New Bold'
			  @ DC_PRINTERROW(),115 DCPRINT SAY NEWUPS->aftsaleam1 PICTURE '99,999.99' FONT '10.Courier New Bold'
		    ENDIF
		ENDIF

		IF NEWUPS->QUOTETYPE $('RC')
			IF lEspanol
			  @ DC_PRINTERROW()+1.1,20 DCPRINT SAY  'Neto de Impuesto'
			  else
			  @ DC_PRINTERROW()+1.1,20 DCPRINT SAY  'Net Taxable'
			endif
		 @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->QUOTED-NEWUPS->TradeQuote PICTURE '999,999.99'
		ELSE
		 @ DC_PRINTERROW()+1.1,20 DCPRINT SAY  'LeaseTermMonths'
		 @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->LTERM PICTURE '999'
		ENDIF
		IF NEWUPS->QUOTETYPE $('RC')
		 IF NEWUPS->REBATE2 > 0
			@ DC_PRINTERROW(),68 DCPRINT SAY NEWUPS->REBDESC2  FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),79 DCPRINT SAY NEWUPS->REBATE2 PICTURE '9,999.99' FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),87 DCPRINT SAY NEWUPS->REBCODE2  FONT '6.Courier New Bold'
		 ENDIF
		 ELSE
		 IF NEWUPS->LREBATE2 > 0
			@ DC_PRINTERROW(),68 DCPRINT SAY NEWUPS->LREBDESC2  FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),79 DCPRINT SAY NEWUPS->LREBATE2 PICTURE '9,999.99' FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),87 DCPRINT SAY NEWUPS->LREBCODE2  FONT '6.Courier New Bold'
		 ENDIF
		ENDIF
		IF lFinal
		 IF ABS(NEWUPS->aftsaleam2) > 0
			@ DC_PRINTERROW(),93 DCPRINT SAY NEWUPS->aftsale2 FONT '8.Courier New Bold'
			@ DC_PRINTERROW(),115 DCPRINT SAY NEWUPS->aftsaleam2 PICTURE '99,999.99' FONT '10.Courier New Bold'
		 ENDIF
		ENDIF
		IF NEWUPS->QUOTETYPE $('RC')
		 IF !lFinal
		  IF lEspanol
			 @ DC_PRINTERROW()+1.2,20 DCPRINT SAY 'Impuesto de Venta @'+alltrim(str(NEWUPS->Taxper))+'%'
			 else
		    @ DC_PRINTERROW()+1.2,20 DCPRINT SAY 'SalesTax @'+alltrim(str(NEWUPS->Taxper))+'%'
		  endif
		  @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->TAX+NEWUPS->AFTSALETAX  PICTURE '999,999.99'
		 else
		  IF lEspanol
			 @ DC_PRINTERROW()+1.2,20 DCPRINT SAY 'Impuesto de Venta @'+alltrim(str(NEWUPS->Taxper))+'%'
			 else
		    @ DC_PRINTERROW()+1.2,20 DCPRINT SAY 'SalesTax @'+alltrim(str(NEWUPS->Taxper))+'%'
		  endif
		  @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->AFTSALETAX+NEWUPS->TAX  PICTURE '999,999.99'
		 ENDIF /// lFinal
		 ELSE
		 @ DC_PRINTERROW()+1.1,20 DCPRINT SAY  'MnthlyLeasePay'
		 @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->LPAYMENT2  PICTURE '999,999.99'
		ENDIF
		IF NEWUPS->QUOTETYPE $('RC')
		 IF NEWUPS->REBATE3 > 0
			@ DC_PRINTERROW(),68 DCPRINT SAY NEWUPS->REBDESC3 FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),79 DCPRINT SAY NEWUPS->REBATE3 PICTURE '9,999.99' FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),87 DCPRINT SAY NEWUPS->REBCODE3  FONT '6.Courier New Bold'
		 ENDIF
		ELSE
		 IF NEWUPS->LREBATE3 > 0
			@ DC_PRINTERROW(),68 DCPRINT SAY NEWUPS->LREBDESC3 FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),79 DCPRINT SAY NEWUPS->LREBATE3 PICTURE '9,999.99' FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),87 DCPRINT SAY NEWUPS->LREBCODE3  FONT '6.Courier New Bold'
		 ENDIF
		ENDIF

		IF lFinal
		 IF ABS(NEWUPS->aftsaleam3) > 0
			@ DC_PRINTERROW(),88 DCPRINT SAY NEWUPS->aftsale3 FONT '8.Courier New Bold'
			@ DC_PRINTERROW(),115 DCPRINT SAY NEWUPS->aftsaleam3 PICTURE '99,999.99' FONT '10.Courier New Bold'
		 ENDIF
		ENDIF

		IF NEWUPS->QUOTETYPE $('RC')
		 IF lEspanol
			@ DC_PRINTERROW()+1.1,15 DCPRINT SAY 'Menos el pago inicial'
			else
		   @ DC_PRINTERROW()+1.1,20 DCPRINT SAY 'Less DownPay'
		 ENDIF
		 @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->DOWNPAY  PICTURE '999,999.99'
		ELSE
		 IF lEspanol
			@ DC_PRINTERROW()+1.1,20 DCPRINT SAY 'Al momento de firmar'
			ELSE
		   @ DC_PRINTERROW()+1.1,20 DCPRINT SAY 'Cash Down'
		 endif
		 @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->DUEATSIGN  PICTURE '999,999.99'
		ENDIF
		IF NEWUPS->QUOTETYPE $('RC')
		 IF NEWUPS->REBATE4 > 0
			@ DC_PRINTERROW(),68 DCPRINT SAY NEWUPS->REBDESC4 FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),79 DCPRINT SAY NEWUPS->REBATE4 PICTURE '9,999.99' FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),87 DCPRINT SAY NEWUPS->REBCODE4  FONT '6.Courier New Bold'
		 ENDIF
		 ELSE
		   IF NEWUPS->LREBATE4 > 0
			 @ DC_PRINTERROW(),68 DCPRINT SAY NEWUPS->LREBDESC4 FONT '6.Courier New Bold'
			 @ DC_PRINTERROW(),79 DCPRINT SAY NEWUPS->LREBATE4 PICTURE '9,999.99' FONT '6.Courier New Bold'
			 @ DC_PRINTERROW(),87 DCPRINT SAY NEWUPS->LREBCODE4  FONT '6.Courier New Bold'
		   ENDIF
		ENDIF


		IF lFinal
		 IF ABS(NEWUPS->aftsaleam4) > 0
			@ DC_PRINTERROW(),88 DCPRINT SAY NEWUPS->aftsale4 FONT '8.Courier New Bold'
			@ DC_PRINTERROW(),115 DCPRINT SAY NEWUPS->aftsaleam4 PICTURE '99,999.99' FONT '10.Courier New Bold'
		 ENDIF
		ENDIF
		IF NEWUPS->QUOTETYPE $('RC')
		 IF lEspanol
			 @ DC_PRINTERROW()+1.2,20 DCPRINT SAY 'Rebaja(s)) '
			 else
		    @ DC_PRINTERROW()+1.2,20 DCPRINT SAY 'Rebate(s) '
		 endif
		 @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->REBATE  PICTURE '999,999.99'
		else
			IF lEspanol
			 @ DC_PRINTERROW()+1.2,20 DCPRINT SAY 'Rebaja(s)) '
			 else
		    @ DC_PRINTERROW()+1.2,20 DCPRINT SAY 'Rebate(s) '
			endif
		 @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->LREBATE  PICTURE '999,999.99'
		ENDIF

		IF NEWUPS->QUOTETYPE $('RC')
		 IF NEWUPS->REBATE5 > 0
			@ DC_PRINTERROW(),68 DCPRINT SAY NEWUPS->REBDESC5 FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),79 DCPRINT SAY NEWUPS->REBATE5 PICTURE '9,999.99' FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),87 DCPRINT SAY NEWUPS->REBCODE5  FONT '6.Courier New Bold'
		 ENDIF
		ELSE
		   IF NEWUPS->LREBATE5 > 0
			 @ DC_PRINTERROW(),68 DCPRINT SAY NEWUPS->LREBDESC5 FONT '6.Courier New Bold'
			 @ DC_PRINTERROW(),79 DCPRINT SAY NEWUPS->LREBATE5 PICTURE '9,999.99' FONT '6.Courier New Bold'
			 @ DC_PRINTERROW(),87 DCPRINT SAY NEWUPS->LREBCODE5  FONT '6.Courier New Bold'
		   ENDIF
		ENDIF
		IF lFinal
		    IF ABS(NEWUPS->aftsaleam5) > 0
			  @ DC_PRINTERROW(),88 DCPRINT SAY NEWUPS->aftsale5 FONT '8.Courier New Bold'
			  @ DC_PRINTERROW(),115 DCPRINT SAY NEWUPS->aftsaleam5 PICTURE '99,999.99' FONT '10.Courier New Bold'
		    ENDIF
		ENDIF
		@ DC_PRINTERROW()+1.2,0 DCPRINT SAY ''
		IF NEWUPS->LIEN <> 0
			IF lEspanol
			  @ DC_PRINTERROW(),20 DCPRINT SAY 'Pago de prestamo'
			  else
			  @ DC_PRINTERROW(),20 DCPRINT SAY 'Loan Pay Off'
			ENDIF
			@ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->LIEN  PICTURE '999,999.99'
		ENDIF
		IF NEWUPS->QUOTETYPE $('RC')
		 IF NEWUPS->REBATE6 > 0
			@ DC_PRINTERROW(),68 DCPRINT SAY NEWUPS->REBDESC6 FONT '6.Courier New Bold'
			@ DC_PRINTERROW(),80 DCPRINT SAY NEWUPS->REBATE6 PICTURE '9,999.99' FONT '6.Courier New Bold'
		 ENDIF
		 ELSE
		   IF NEWUPS->LREBATE6 > 0
			 @ DC_PRINTERROW(),68 DCPRINT SAY NEWUPS->LREBDESC6 FONT '6.Courier New Bold'
			 @ DC_PRINTERROW(),79 DCPRINT SAY NEWUPS->LREBATE6 PICTURE '9,999.99' FONT '6.Courier New Bold'
			 @ DC_PRINTERROW(),87 DCPRINT SAY NEWUPS->LREBCODE6  FONT '6.Courier New Bold'
		   ENDIF
		ENDIF

		IF lFinal
		 IF ABS(NEWUPS->aftsaleam6) > 0
			@ DC_PRINTERROW(),88 DCPRINT SAY NEWUPS->aftsale6 FONT '8.Courier New Bold'
			@ DC_PRINTERROW(),115 DCPRINT SAY NEWUPS->aftsaleam6 PICTURE '99,999.99' FONT '10.Courier New Bold'
		 ENDIF
		ENDIF
		//IF NEWUPS->EXTRANUM > .01
		// @ DC_PRINTERROW()+1.1,20 DCPRINT SAY NEWUPS->EXTRACHAR
		// @ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->EXTRANUM PICTURE  '999,999.99'
		//ENDIF
		IF NEWUPS->QUOTETYPE $('RC')
			IF lEspanol
				@ DC_PRINTERROW()+1.1,2 DCPRINT SAY 'Cantidad debido del vehiculo'
				else
		      @ DC_PRINTERROW()+1.1,15 DCPRINT SAY 'Amount Due Vehicle'
			ENDIF
		@ DC_PRINTERROW(),48 DCPRINT SAY (NEWUPS->QUOTED+NEWUPS->TAX+NEWUPS->AFTSALETAX+NEWUPS->LIEN);
			-(NEWUPS->REBATE+NEWUPS->DOWNPAY+NEWUPS->TradeQuote)  PICTURE '999,999.99'
	else
		IF lEspanol
		  @ DC_PRINTERROW()+1.1,15 DCPRINT SAY 'Millas por Ano'
		  else
		  @ DC_PRINTERROW()+1.1,20 DCPRINT SAY 'MilesPerYear'
		endif
		@ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->LMILES PICTURE '999999'
		ENDIF
		IF lFinal
		 IF NEWUPS->ServConAm > 0
			 @ DC_PRINTERROW(),75 DCPRINT SAY NEWUPS->servcont FONT '8.Courier New Bold'
			 @ DC_PRINTERROW(),115 DCPRINT SAY NEWUPS->servconam PICTURE '99,999.99' FONT '10.Courier New Bold'
	  	 ENDIF
		ENDIF


		//IF NEWUPS->AFTSALETAX > 0
		 //	@ DC_PRINTERROW(),88 DCPRINT SAY 'Added Sales Tax' FONT '6.Courier New Bold'
		 //	@ DC_PRINTERROW(),116 DCPRINT SAY NEWUPS->aftsaletax PICTURE '9,999.99' FONT '10.Courier New Bold'
		//ENDIF

		  /// now print comments
		IF !empty(NEWUPS->BoComment)
			cMEMO:=ALLTRIM(NEWUPS->BOCOMMENT )
		   nLINECOUNT:=mlcOUNT(cMEMO)
		   FOR I:=1 TO nLINECOUNT+1
			cMEMOLINE:=MEMOLINE(cMEMO,60,I)
			@ DC_PRINTERROW()+1.1,5 DCPRINT SAY cMEMOLINE FONT '10.Times New Roman Bold'
			NEXT
		ENDIF

		IF lFinal
		 @ 37,111 DCPRINT SAY NEWUPS->QUOTED+NEWUPS->AFTSALEAM1+NEWUPS->AFTSALEAM2+NEWUPS->AFTSALEAM3+NEWUPS->AFTSALEAM4+NEWUPS->AFTSALEAM5+;
			NEWUPS->AFTSALEAM6+NEWUPS->SERVCONAM PICTURE '999,999.99'
		ENDIF

		/// trade section
		@ 38,66 DCPRINT SAY NEWUPS->INSURANCE FONT '8.Tahoma Bold'

		@ 39,111 DCPRINT SAY NEWUPS->TRADEQUOTE PICTURE '999,999.99'
		@ 40.5,5 DCPRINT SAY NEWUPS->TRADEYR
		@ DC_PRINTERROW(),16 DCPRINT SAY NEWUPS->TRADEMILES PICTURE '999999'
		@ DC_PRINTERROW(),30 DCPRINT SAY NEWUPS->TRADEMAKE  FONT '10.COURIER NEW BOLD'
		@ DC_PRINTERROW(),49 DCPRINT SAY SUBSTR(NEWUPS->TRADEDESC,1,10)  FONT '10.Courier New Bold'
		@ DC_PRINTERROW(),64 DCPRINT SAY NEWUPS->TRADECOLOR

		@ DC_PRINTERROW()+1.5,33 DCPRINT SAY NEWUPS->TRADEVIN
		IF NEWUPS->QUOTETYPE='L'
			IF lEspanol
			 @ DC_PRINTERROW(),85 DCPRINT SAY 'AL MOMENTO DE FIRMAR'
		   else
			 @ DC_PRINTERROW(),85 DCPRINT SAY 'DUE AT SIGNING'
			endif
		ENDIF
		@ DC_PRINTERROW()+1.2,5 DCPRINT SAY SUBSTR(NEWUPS->LIENINFO,1,40) FONT '8.Courier New'
		@ DC_PRINTERROW(),59 DCPRINT SAY NEWUPS->LIEN  picture '99,999.99' FONT '10.Courier New Bold'
		IF lFinal
			IF NEWUPS->QUOTETYPE $('RC')
			@ DC_PRINTERROW(),111 DCPRINT SAY (NEWUPS->QUOTED)-(NEWUPS->TradeQuote)+;
				(NEWUPS->AFTSALEAM1+NEWUPS->AFTSALEAM2+NEWUPS->AFTSALEAM3+NEWUPS->AFTSALEAM4+NEWUPS->AFTSALEAM5+NEWUPS->AFTSALEAM6+;
				 NEWUPS->SERVCONAM) PICTURE '999,999.99'
			ELSE
			@ DC_PRINTERROW(),111 DCPRINT SAY NEWUPS->DUEATSIGN PICTURE '999,999.99'
			ENDIF

			IF NEWUPS->QUOTETYPE $('RC')
			 @ DC_PRINTERROW()+2.4,78 DCPRINT SAY alltrim(NEWUPS->COUNTY)+' '+alltrim(str(NEWUPS->Taxper))+'%'
			 @ DC_PRINTERROW(),111 DCPRINT SAY NEWUPS->aftsaletax+NEWUPS->TAX PICTURE '999,999.99'
			ELSE
			 @ DC_PRINTERROW()+2.4,78 DCPRINT SAY ''
			ENDIF
			@ DC_PRINTERROW()+1,111 DCPRINT SAY NEWUPS->PROCFEE  PICTURE '999,999.99'
			IF NEWUPS->REGTRANS='Y'
			 @ DC_PRINTERROW()+1,90 DCPRINT SAY 'TRANSFER'
			ELSE
				@ DC_PRINTERROW()+1,90 DCPRINT SAY 'NEW'
			ENDIF
			@ DC_PRINTERROW(),111 DCPRINT SAY NEWUPS->REGFEE   PICTURE '999,999.99'
			@ DC_PRINTERROW()+1,111 DCPRINT SAY NEWUPS->INSPECTION PICTURE '999,999.99'
			@ DC_PRINTERROW()+1,111 DCPRINT SAY NEWUPS->WASTEFEE   PICTURE '999,999.99'
			IF NEWUPS->QUOTETYPE $('RC')
			   @ DC_PRINTERROW()+1.2,111 DCPRINT SAY (NEWUPS->QUOTED+NEWUPS->TAX)-(NEWUPS->TradeQuote)+;
				(NEWUPS->AFTSALEAM1+NEWUPS->AFTSALEAM2+NEWUPS->AFTSALEAM3+NEWUPS->AFTSALEAM4+NEWUPS->AFTSALEAM5+NEWUPS->AFTSALEAM6+NEWUPS->SERVCONAM+;
				 NEWUPS->AFTSALETAX+NEWUPS->PROCFEE+NEWUPS->REGFEE+NEWUPS->INSPECTION+NEWUPS->WASTEFEE)  PICTURE '999,999.99'
			ELSE

				@ DC_PRINTERROW()+1.2,111 DCPRINT SAY NEWUPS->DUEATSIGN PICTURE '999,999.99' //  +NEWUPS->PROCFEE+NEWUPS->REGFEE+NEWUPS->INSPECTION+;
					//NEWUPS->WASTEFEE)-(NEWUPS->Tradequote-NEWUPS->LIEN) PICTURE '999,999.99'
			ENDIF
			IF NEWUPS->QUOTETYPE='L'
			 @ DC_PRINTERROW()+1.2,85 DCPRINT SAY 'IncludedAbove'
			ELSE
			 @ DC_PRINTERROW()+1.2,90 DCPRINT SAY ''
			ENDIF
		   @ DC_PRINTERROW(),111 DCPRINT SAY NEWUPS->REBATE   PICTURE '999,999.99'
			@ DC_PRINTERROW()+1,111 DCPRINT SAY NEWUPS->DEPOSIT  PICTURE '999,999.99'

			IF NEWUPS->QUOTETYPE='L'
			 @ DC_PRINTERROW()+1.2,85 DCPRINT SAY 'IncludedAbove'
			ELSE
			 @ DC_PRINTERROW()+1.2,90 DCPRINT SAY ''
			ENDIF

			@ DC_PRINTERROW(),111 DCPRINT SAY NEWUPS->LIEN     PICTURE '999,999.99'
			IF NEWUPS->QUOTETYPE $('RC')
			 @ DC_PRINTERROW()+1,111 DCPRINT SAY	NEWUPS->FINANCE  PICTURE '999,999.99'
			ELSE
			 @ DC_PRINTERROW()+1,111 DCPRINT SAY ''
			ENDIF
			IF NEWUPS->QUOTETYPE $('RC')
		   	@ DC_PRINTERROW()+1,111 DCPRINT SAY  NEWUPS->COD PICTURE '999,999.99'
			ELSE
				@ DC_PRINTERROW()+1,111 DCPRINT SAY NEWUPS->DUEATSIGN-NEWUPS->DEPOSIT PICTURE '999,999.99'
               //+NEWUPS->PROCFEE+NEWUPS->REGFEE+NEWUPS->INSPECTION+;
				  //	NEWUPS->WASTEFEE)-(NEWUPS->DEPOSIT+(NEWUPS->Tradequote-NEWUPS->LIEN)) PICTURE '999,999.99'
			ENDIF
		 ELSE
			@ DC_PRINTERROW()+3.4,111 DCPRINT SAY NEWUPS->PROCFEE  PICTURE '999,999.99'
			IF NEWUPS->NEWUSED='N'
			 @ DC_PRINTERROW()+3,111 DCPRINT SAY NEWUPS->WASTEFEE   PICTURE '999,999.99'
			ENDIF
		ENDIF
		IF nNUMCOPY=2
		 @ 65,40 DCPRINT SAY 'Customer Copy'
		 DCPRINT EJECT
		ELSE
		 @ 65,40 DCPRINT SAY 'Dealership Copy'
		ENDIF

		nNumcopy:=nNumCopy-1
   enddo  /// nNumcopy

	IF !lFinal
	 IF lPrintSupplement
	  IF FEXISTS(GETFLG('ACCTPATH')+'DELVLIST.JPG')
	   DCPRINT EJECT
	   @ 0,0,66,133 dcprint bitmap (GETFLG('ACCTPATH')+'delvlist.bmp')
	   @ 2,10 DCPRINT SAY DTOC(DATE())+' '+TIME()
	   @ DC_PRINTERROW()+1,10 DCPRINT SAY ALLTRIM(NEWUPS->LASTNAME)+','+ALLTRIM(NEWUPS->FIRSTNAME)
	   @ DC_PRINTERROW(),102  DCPRINT SAY NEWUPS->SALESMAN+' '+cSmName
	  //DCPRINT OFF
	  ENDIF
	 ENDIF
	endif

	IF !lPrintTerms
		DCPRINT OFF
   ELSE
	   DCPRINT EJECT
	ENDIF

	IF lPrintTerms
		IF lEspanol
		 @ 0,0,66,133 dcprint bitmap (GETFLG('ACCTPATH')+'buyordtermspanish.jpg')
		 @ 1,10 DCPRINT SAY DTOC(DATE())+' '+TIME()
	    @ DC_PRINTERROW()+1,10 DCPRINT SAY ALLTRIM(NEWUPS->LASTNAME)+','+ALLTRIM(NEWUPS->FIRSTNAME)
		 @ DC_PRINTERROW(),102  DCPRINT SAY NEWUPS->SALESMAN+' '+cSmName
		else
		 @ 0,0,66,133 dcprint bitmap (GETFLG('ACCTPATH')+'buyersordterms.jpg')
		 @ 1,10 DCPRINT SAY DTOC(DATE())+' '+TIME()
	    @ DC_PRINTERROW()+1,10 DCPRINT SAY ALLTRIM(NEWUPS->LASTNAME)+','+ALLTRIM(NEWUPS->FIRSTNAME)
		 @ DC_PRINTERROW(),102  DCPRINT SAY NEWUPS->SALESMAN+' '+cSmName
		 endif
		 DCPRINT OFF
	ENDIF


	// if this is just a pdf send an email and return
	if lPDF
		BOPDFemail(NEWUPS->email)
		return nil
	endif

	oRecord:WRITEUP:=.T.

	RETURN NIL




/// retrieve sales person name
function getsmname(cSM,cSmName)
select SM
seek cSM
IF found()
	cSmName:=SM->Name
ENDIF
return cSmName

function getsmnameCA(cSM,cSmFirst,cSMLast)
local nBreak:=0
select SM
seek cSM
IF !found()
	DC_WINALERT('Person Not Found')
ENDIF
IF found()
	nBreak:=AT(' ',SM->INETNAME)
	cSmFirst:=substr(SM->INETNAME,1,nBreak)
	cSmLast:=substr(SM->INETNAME,(nBreak+1),(20-len(cSmFirst)))
ENDIF
return nil


/// retrieve vehicle info
static function GETNEWINFO(cStockno,cMake,cModel,cSeries,cVehicleId,cYear,cColor, nMileage,cInterior,cType)
select nci
seek cStockno
IF found()
	cMake:=NCI->Make
	cModel:=NCI->Model
	cSeries:=NCI->TRIM
	cVehicleID:=NCI->VIN
	cYear:=NCI->Year
	cColor:=NCI->FactColor
	nMileage:=NCI->Mileage
	cInterior:=NCI->Intcode
	cType:=NCI->Body
endif
return nil

/// retrieve vehicle info
static function GETUSEDINFO(cStockno,cMake,cModel,cSeries,cVehicleId,cYear,cColor, nMileage,cInterior,cType,cFmrLR)
select UCL01
seek cStockno
IF found()
	cMake:=UCL01->Make
	cModel:=UCL01->Model
	cSeries:=UCL01->TRIM
	cVehicleID:=UCL01->VIN
	cYear:=UCL01->Year
	cColor:=Ucl01->Color
	nMileage:=UCL01->Mileage
	cInterior:=UCL01->IntColor
	cType:=UCL01->Body
	cFmrLR:=UCL01->FMR_LR
endif
return nil

// THIS IS THE BROWSER BY SALESMAN NUMBER
FUNCTION BROWBYSM()
	LOCAL GETLIST:={}  , oBrowse , Aotoolbar  , aPres  , oRecord
	LOCAL cSM:=SPACE(3), cProgram , getoptions, xStatus , aRecords:={}
	LOCAL dDatex:=(DATE()-10)   , lNoShowDead:=.t.
	cProgram:='BROWBYSM'
	IF.NOT.SECURITY(@cProgram)
	 DC_WINALERT('SECURITY VIOLATION')
    RETURN NIL
   ENDIF

   IF !OPENUPSFILES()
	 CLOSE ALL
    RETURN NIL
   ENDIF



NEWUPS->(ORDSETFOCUS('NEWUPSM'))


DCGETOPTIONS ;
	SAYWIDTH 0 ;
	AUTORESIZE
XSTATUS:=.T.
dDatex:=DATE()-60
@ 10,0 DCSAY 'ENTER THE SALESPERSON TO BROWSE' GET cSM SAYCOLOR 3,0 SAYFONT '10.COURIER' GETFONT '12.COURIER' ;
	picture '!!'
@ 12,0 DCSAY 'ENTER THE DATE TO GO BACK TO BEGIN BROWSE' GET dDatex SAYCOLOR 9,0 SAYFONT '10.COURIER' GETFONT '12.COURIER'
//@ 14,0 DCCHECKBOX lNoShowDead PROMPT 'Click here to NOT show deals already DEAD'

DCREAD GUI TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'ENTER LASTNAME TO BEGIN BROWSE'
IF !XSTATUS
	CLOSE DATABASES
RETURN NIL
ENDIF
SELECT NEWUPS
ordsetfocus('newupsm')
SEEK cSM

IF lNoShowDead
	SET FILTER TO NEWUPS->upstatus <> 'D'
ENDIF




/// now get to date requested
do while NEWUPS->SALESMAN=cSm
	IF NEWUPS->datein >=dDatex
		AADD(aRecords,NEWUPS->(RECNO()) )
	ENDIF
	skip alias newups
enddo

/// now create browse


DC_SetScopeArray(aRecords)
	DC_DbGoTop()

aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE }, ;//     Header FG Color       ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY },; //  Header BG Color       ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },;   //Row Sep       ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED }, ;  //Col Sep       ;
    { XBP_PP_COL_DA_ROWHEIGHT, 45 },                ; //  Row Height      ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, BD_PALEBLUE },;  // HILITE BG Color   ;
	 { XBP_PP_COL_DA_HILITE_FGCLR, GRA_CLR_BLACK},;  // HILITE BG Color   ;
    { XBP_PP_COL_DA_CELLHEIGHT, 10 }  }              // Cell Height

/* ----- Create ToolBar ----- */

@ 1,1 DCTOOLBAR AoToolBar                               ;
	SIZE 125,2  BUTTONSIZE 15,2


DCADDBUTTON CAPTION 'New Vehicle;Browser  '                              ;
	ACTION {||SBROWBYDESC(),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar

DCADDBUTTON CAPTION 'Used Vehicle;Browser '                              ;
	ACTION {||SBROWBYDESC(),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar

DCADDBUTTON CAPTION 'Maintain Up;Record '                              ;
	ACTION {||UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar


DCADDBUTTON CAPTION 'COMMENTS '                              ;
	ACTION {|| UPNOTES(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(obROWSE:GETCOLUMN(1))  };
	PARENT AoToolBar

DCADDBUTTON CAPTION 'Print UPSheet '                              ;
	ACTION {||PRINTUPSHEET(NEWUPS->(RECNO()))};
	PARENT aoToolBar

DCADDBUTTON CAPTION 'Send eMail '                              ;
	ACTION {||upemail()};
	PARENT aoToolBar

DCADDBUTTON CAPTION 'Review eMails '                              ;
		ACTION {|| BROWEMAILLOG('',NEWUPS->EMAIL)} ;
		PARENT AoToolBar


@ 4,0 DCSAY 'COLOR CODES'
@ 4,15 DCSAY 'Sold  ' SAYCOLOR 1,BD_KEYLIME
@ 4,30 DCSAY 'Active' SAYCOLOR 1, BD_WASHGREY
@ 4,45 DCSAY 'Tickle' SAYCOLOR 1,BD_CORNFLOWERBLUE
@ 4,60 DCSAY 'DEAD  ' SAYCOLOR 1,BD_SALMON
@ 4,75 DCSAY 'Rescue' SAYCOLOR 2,BD_SUNNYYELLOW
@ 4,90 DCSAY 'PhoneUp' SAYCOLOR 2,BD_FADEDPURPLE
@ 4,105 DCSAY 'InterNetUp' SAYCOLOR 1,BD_BLUEMIST



/* ----- Create browse ----- */

@ 6,1 DCBROWSE oBrowse ALIAS 'NEWUPS'                 ;
	SIZE 170,20                                     ;
	PRESENTATION aPres;
	FONT '8.Lucida Console'  ;
	HEADLINES 4 ;//FREEZELEFT {1,2,3} ;
	SCOPE;
	DATALINK{|| UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST) };

DCBROWSECOL FIELD NEWUPS->SALESMAN ;
	HEADER "**SALESMAN" PARENT oBrowse  ;
	WIDTH 3

DCBROWSECOL FIELD NEWUPS->DATEIN ;
	HEADER "**DATEIN" PARENT oBrowse  ;
	WIDTH 8


DCBROWSECOL FIELD NEWUPS->UPSTATUS ;
	HEADER "STATUS" PARENT oBrowse COLOR {|| getupcolor(NEWUPS->UPSTATUS)} ;
	WIDTH 1

DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "**UPID/PHN#" PARENT oBrowse COLOR {|| getupcolor(NEWUPS->UPSTATUS)};
	WIDTH 7


DCBROWSECOL DATA {{|| NEWUPS->LASTNAME},;
		               {|| NEWUPS->FIRSTNAME},;
							{|| NEWUPS->STREET },;
							{|| NEWUPS->CITYSTATE+' '+NEWUPS->ZIPCODE}};
	WIDTH 25 HEADER "LastName;FirstName;Street;CitySt" PARENT oBrowse ;
	EDITPROTECT{|| TRUE }


DCBROWSECOL DATA {{|| TRANSFORM(NEWUPS->AREA+NEWUPS->UPID,'@R 999.999.9999')},;
						{|| TRANSFORM(NEWUPS->WORKPHN,'@R 999.999.9999.99999')},;
						{|| TRANSFORM(NEWUPS->CELLPHONE,'@R 999.999.9999')},;
						{|| NEWUPS->email}};
	HEADER 'Phone;WorkPhn;CellPhn;eMail';
	width 25 PARENT oBrowse ;
	EDITPROTECT{|| TRUE }




DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "**TYPE" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1

DCBROWSECOL FIELD NEWUPS->ORIGSOURCE ;
	HEADER "**Source" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1




DCBROWSECOL DATA { ;
                 {|| NEWUPS->BEBACK1},;
	              {|| NEWUPS->BEBACK2}};
	HEADER "BeBack1;BeBack2" PARENT oBrowse  ;
	WIDTH 8

DCBROWSECOL DATA { ;
						{|| IIF(NEWUPS->NEWUSED='N' ,'NEW' ,'USED' )},;
	               {|| NEWUPS->INTEREST}};
	HEADER "New/Used;Interest" PARENT oBrowse  ;
	WIDTH 10

DCBROWSECOL DATA { ;
						{|| NEWUPS->ADSOURCE},;
	               {|| NEWUPS->TOMGR}};
	HEADER "Source/TOMgr;Interest" PARENT oBrowse  ;
	WIDTH 8


DCBROWSECOL DATA { ;
						{|| NEWUPS->COMMENTS},;
	               {|| NEWUPS->COMMENT2}};
	HEADER "Comments" PARENT oBrowse  ;
	WIDTH 20



/*
DCBROWSECOL FIELD NEWUPS->Upstatus ;
	HEADER "Status" PARENT oBrowse Hcolor 15,10  ;
	WIDTH 3;
	VALID {|| IIF(NEWUPS->UPSTATUS $('ARFDS'),.T.,.F.)}

DCBROWSECOL FIELD NEWUPS->Billstat ;
	HEADER "BillStat" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 4 ;
	VALID {|| IIF(NEWUPS->BILLSTAT $(' ID'),.t.,.f.)}

DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->ORIGSOURCE ;
	HEADER "Source" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1

DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "UPID/PHN#" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 7 ;
	VALID{|| IIF(EMPTY(NEWUPS->UPID),.F.,.T.)}

DCBROWSECOL FIELD NEWUPS->PREFMETHOD ;
	HEADER "PREFMETHOD" PARENT oBrowse  ;
	WIDTH 4



DCBROWSECOL FIELD NEWUPS->SALESMAN ;
	HEADER "SALESMAN" PARENT oBrowse Hcolor 15,10 ;
	WIDTH 3 ;
	Editprotect{|| .t.}

DCBROWSECOL FIELD NEWUPS->TOMgr ;
	HEADER "T.O.Mgr" PARENT oBrowse Hcolor 15,10 ;
	WIDTH 3 ;
	Editprotect{|| .t.}

DCBROWSECOL FIELD NEWUPS->Causedead ;
	HEADER "COD" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 1 ;
	VALID {|| IIF(NEWUPS->Causedead $('CBIONU'),.t.,.f.)}

DCBROWSECOL FIELD NEWUPS->LASTNAME                     ;
	HEADER "LASTNAME" HCOLOR 0,12 PARENT oBrowse   ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->FIRSTNAME ;
	HEADER "FIRSTNAME" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->DATEIN ;
	HEADER "DATEIN" PARENT oBrowse Hcolor 15,10 ;
	WIDTH 8;
	Editprotect{|| .t.}
DCBROWSECOL FIELD NEWUPS->EMAIL ;
	HEADER "EMAIL" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 20


DCBROWSECOL FIELD NEWUPS->STREET ;
	HEADER "STREET" PARENT oBrowse  Hcolor 0,12;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->CITYSTATE ;
	HEADER "CITYSTATE" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->ZIPCODE ;
	HEADER "ZIPCODE" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 5
DCBROWSECOL FIELD NEWUPS->AREA ;
	HEADER "AREACODE" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "PHONE" PARENT oBrowse  Hcolor 0,12;
	WIDTH 7
DCBROWSECOL FIELD NEWUPS->WORKPHN ;
	HEADER "WORKPHN" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->cellphone ;
	HEADER "CELLPHN" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 10

DCBROWSECOL FIELD NEWUPS->EMAIL ;
	HEADER "EMAIL" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 35
DCBROWSECOL FIELD NEWUPS->NEWUSED ;
	HEADER "NEWUSED" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 1

DCBROWSECOL FIELD NEWUPS->Tickle ;
	HEADER "Follow up date" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 10

DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->INTEREST ;
	HEADER "INTEREST" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->SALECODE ;
	HEADER "SALECODE" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->COMMENTS ;
	HEADER "COMMENT" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->COMMENT2 ;
	HEADER "COMMENT2" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->ADSOURCE ;
	HEADER "ADSOURCE" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 4
DCBROWSECOL FIELD NEWUPS->DOWNDEAL ;
	HEADER "DOWN" PARENT oBrowse Hcolor 0,12 ;
	WIDTH 1   */



DCREAD GUI ;
	FIT ;
	BUTTONS DCGUI_BUTTON_EXIT   ;
	OPTIONS GETOPTIONS ;
	title 'Automan Up Browser by SalesPerson'
CLOSE ALL
ReTURN nil

static function getupcolor(uPstat)
local getlist:={}
@ 3,15 DCSAY 'Sold' SAYCOLOR 1,BD_KEYLIME
@ 3,30 DCSAY 'Active' SAYCOLOR 1, BD_WASHGREY
@ 3,45 DCSAY 'Tickle' SAYCOLOR 1,BD_CORNFLOWERBLUE
@ 3,60 DCSAY 'DEAD' SAYCOLOR 1,BD_SALMON
@ 3,75 DCSAY 'Rescue' SAYCOLOR 2,BD_SUNNYYELLOW
@ 3,90 DCSAY 'PhoneUp' SAYCOLOR 2,BD_FADEDPURPLE
@ 3,105 DCSAY 'InterNetUp' SAYCOLOR 1,BD_BLUEMIST



	if upstat='R'
		return {GRA_CLR_BLUE,BD_SUNNYYELLOW}
	ENDIF
	if upstat='D'
		return {GRA_CLR_BLACK,BD_SALMON}
	ENDIF
	if upstat='F'
		return {GRA_CLR_BLACK,BD_CORNFLOWERBLUE}
	ENDIF
	if upstat='S'
		return {GRA_CLR_BLACK,BD_KEYLIME}
	ENDIF

RETURN   {GRA_CLR_BLACK,BD_WASHGREY}

static function getphoneupcolor(uPSRC)

if uPSRC='P'
		return {GRA_CLR_BLUE,BD_FADEDPURPLE}
ENDIF

if uPSRC='I'
		return {GRA_CLR_BLACK,BD_BLUEMIST}
ENDIF


RETURN   {GRA_CLR_BLACK,GRA_CLR_WHITE}



FUNCTION SALEFOLO()
	local getlist:={},nPage
	local getoptions , xStatus , clSorton , dBegdt, dEnddt, cSM, cSaletype , lPrevue, cprogram ,oPrinter
	LOCAL SYS_TITLE,PRG_NAME
	TOP_MAR(3)
	BOT_MAR(2)
	PAGE_LEN(62)
	nPage:=0
	SYS_TITLE:='DELIVERED SALES FOLLOW UP REPORT'
	PRG_NAME:='SALEFOLO'
	clSORTON:='Y'
	@ 2,0 DCSAY 'DO YOU WISH TO CREATE NEW SORT FILE Y/N' GET clSORTON PICTURE '@! L'
	DCREAD GUI APPWINDOW MAINWINDOW():DRAWINGAREA fit TO XSTATUS ENTEREXIT title 'Create New Sort'
	if !xstatus
		return nil
	endif


	IF !UseDb({'SERVCUST'})
		 DC_WINALERT('Cannot Open Files     .')
	   	 RETURN NIL
	ENDIF
	SERVCUST->(ORDSETFOCUS('SERVSM'))
	SELECT SERVCUST
	dBEGDT:=CTOD("  /  /  ")
	dENDDT:=CTOD("  /  /  ")
	cSALETYPE:='N'
	cSM:='  '
	@ 4,0 DCSAY 'ENTER BEGINNING DELIVERY DATE ' GET dBEGDT
	@ 5,0 DCSAY 'ENTER ENDING DELIVERY DATE' GET dENDDT
	@ 6,0 DCSAY 'ENTER SALESMANS INITIALS' GET cSM VALID(SMIGET(cSM))
	@ 7,0 DCSAY 'ENTER SALE TYPE N FOR NEW U FOR USED' GET cSALETYPE
	DCREAD GUI APPWINDOW MAINWINDOW():DRAWINGAREA fit TO XSTATUS ENTEREXIT title 'Create New Sort'
	lPrevue:=.f.
	TESTFORPREV(@lPrevue)
	DCPRINT ON SIZE 66,132 TO oPrinter FONT '8.Courier New' CANCELENABLE _PREVIEW lPrevue zoomfactor 2.0
	IF VALTYPE(OPRINTER) # 'O' .OR. !OPRINTER:LACTIVE

		RETURN NIL
	ENDIF

	@ dc_printerrow()+1,dc_printerCOL() DCPRINT SAY SYS_TITLE
	@ dc_printerrow(),dc_printerCOL()+4 DCPRINT SAY DATE()
	@ dc_printerrow(),dc_printerCOL()+4 DCPRINT SAY TIME()
	@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY PRG_NAME

	BEGIN SEQUENCE
		SELECT SERVCUST
		SEEK cSM

		CUSSTHDR(@nPage)
		DO WHILE SERVCUST->SM=cSM
			IF DCPAGEEJECT()
				CUSSTHDR()
			ENDIF
			IF SERVCUST->SALETYPE=cSALETYPE
				IF SERVCUST->INVDATE >=dBEGDT.AND.SERVCUST->INVDATE <=dENDDT
					@dc_printerrow()+1,0 DCPRINT SAY SERVCUST->SM
					@dc_printerrow(),dc_printercOL()+2 DCPRINT SAY SERVCUST->LASTNAME
					@dc_printerrow(),dc_printerCOL()+2 DCPRINT SAY SERVCUST->FIRSTNAME
					@dc_printerrow(),dc_printercOL()+2 DCPRINT SAY SERVCUST->CARDESC
					IF DCPAGEEJECT()
						CUSSTHDR(@nPage)
					ENDIF

					@dc_printerrow()+1,0 DCPRINT SAY SERVCUST->STREET
					@dc_printerrow(),dc_printercOL()+2 DCPRINT SAY SERVCUST->CITYSTATE
					@dc_printerrow(),dc_printercOL()+2 DCPRINT SAY SERVCUST->ZIP
					@dc_printerrow(),dc_printercOL()+2 DCPRINT SAY SERVCUST->PHONE  PICTURE '9999999999'

					IF DCPAGEEJECT()
						CUSSTHDR(@nPage)
					ENDIF

					@dc_printerrow()+1,0 DCPRINT SAY SERVCUST->SALETYPE
					@dc_printerrow(),dc_printercOL()+1 DCPRINT SAY SERVCUST->WPHONE
					@dc_printerrow(),dc_printercOL()+2 DCPRINT SAY SERVCUST->INVDATE
					@dc_printerrow(),dc_printercOL()+2 DCPRINT SAY SERVCUST->LASTSERV
					@dc_printerrow(),dc_printercOL()+2 DCPRINT SAY SERVCUST->LASTMILES
					IF DCPAGEEJECT()
						CUSSTHDR(@nPage)
					ENDIF

					@ dc_printerrow()+2,0 DCPRINT SAY ''
				ENDIF
			ENDIF
			SKIP ALIAS SERVCUST
		ENDDO
	END SEQUENCE
	dcprint off

	CLOSE DATABASES
RETURN .T.

// END SALEFOLO
FUNCTION SMIGET(SMI)
	IF SMI='  '
		RETURN .F.
	ENDIF
RETURN .T.

FUNCTION CUSSTHDR(nPage)
	nPage++
	@ dc_printerrow()+1,0 SAY 'SM / LASTNAME / FIRSTNAME / VEHICLE SOLD /     Page' + alltrim(str(nPage))
	@ dc_printerrow()+1,0 SAY 'STREET ADDRESS   /   CITY STATE      ZIP CODE  HOME PHONE'
	@ dc_printerrow()+1,0 SAY 'SALE TYPE /WORK PHONE  / DELIVERY DATE / LAST SERVICED / MILES'
RETURN .T.


FUNCTION CHKSM(SAL)
	SELECT SM
	SEEK SAL
	IF FOUND()
		RETURN .T.
	ELSE
		RETURN .F.
	ENDIF
RETURN .F.


FUNCTION CHKNU(NEW)
	IF NEW $('NU')
		RETURN .T.
	ENDIF
RETURN .F.


*UPENTRY
FUNCTION UPENTRY()
	LOCAL GETLIST:={}
	LOCAL MCON:='Y' ,aPres:={}  , oBrowse
	LOCAL GETOPTIONS , cProgram,  xStatus ,cFranchise, cUpMessage:=space(60)
	local oRecord,lBeback:=.f.
	local lSearchserv:=.t., nRec:=0
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE SAYFONT '12.COURIER NEW' GETFONT '12.COURIER NEW' EXITVALIDATE
	cProgram:='UPENTRY'
	IF.NOT.SECURITY(@cProgram)
		DC_WINALERT('SECURITY VIOLATION')
	 RETURN NIL
	ENDIF


 	 aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_ROWHEIGHT, -1 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 15 }  }              /* Cell Height */



IF !OPENUPSFILES()
	CLOSE ALL
   RETURN NIL
ENDIF

NEWUPS->(ORDSETFOCUS('NEWUPIN'))
INETUPS->(ORDSETFOCUS('INETDATE'))
INETUPS->(DBSEEK((DATE()-1),.T.,'INETDATE',.T.))    /// GET LAST UP FROM YESTERDAY TO START DISPLAY

IF !UseDb({'COMPANY','HNEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF

cFranchise:=COMPANY->FRANCHISE

oRecord:=NEWUPS->(DC_DBRECORD(): NEW())
NEWUPS->(DB_INIT(oRecord))
oRecord:DATEIN:=DATE()
// initialize lease captax to Y


oRecord:captax:='Y'
oRecord:capfees:='N'

/// gather initial up info

	@ 1,20 DCSAY 'I n i t i a l U P  E n t r y  S y s t e m '

   @ 2,5 DCSAY 'To Record and Save a New Up You Must At Least Enter Sales Person ID and Customer Id and Last Name On This Screen' SAYCOLOR 1,BD_SUNNYYELLOW
	@ 3,5 DCSAY '      Enter Date For Ups Being Recorded Or Esc To Exit'  GET oRecord:DATEIN
	@ 4,5 DCSAY '                        Enter Valid Salesmans Initials' Get oRecord:SALESMAN Valid{|| Chksm(oRecord:SALESMAN)} ;
		getcolor{|| IIF(EMPTY(oRecord:salesman),{1,BD_SALMON} ,{1,BD_KEYLIME}  ) }   Picture '@!'

  	@ 5,5 DCSAY ' Customer Id- Enter 7 Digit Phone# Or First 7 Of eMail' GET oRecord:UPID;
		Saycolor{|| IIF(EMPTY(oRecord:UPID),{1,BD_SALMON} ,{1,BD_KEYLIME}  )}    Picture '@!' ;
		valid { || _lookupup(@oRecord,lSearchserv,@cUpMessage,@lBeback),DC_GETREFRESH(GETLIST),.T.} ;    /// send lsearchserv to launch serv lookup or not
		POPUP{|| UPSEARCHBUTTON(@oRecord,@lSearchserv),DC_GETREFRESH(GETLIST)};  ///send and modify lsearchserv
		POPCAPTION 'Search Service Customers ' POPWIDTH 190 POPFONT '8.Courier New' POPSTYLE 1
   @  6,5 DCSAY '                                    Customer Last Name' GET oRecord:LASTNAME  getcolor{|| IIF(EMPTY(oRecord:LASTNAME),{1,BD_SALMON} ,{1,BD_KEYLIME}  ) }
   @  7,5 DCSAY '                                   Customer First Name' GET oRecord:FIRSTNAME

 /*
  @ 16,5 DCPUSHBUTTON CAPTION 'Search InterNet Leads' size 30,2 ;
		CARGO 'CANCEL' ;
		WHEN {|| IIF( !EMPTY(oRecord:salesman),.t. ,.f. )};
		ACTION {|| _searchinetups(@oRecord),DC_GETREFRESH(GETLIST)}      */


@ 10,0 DCSAY 'Latest Internet Ups to Select From. You May Double Click On Record To Select '
  @ 11,1 DCBROWSE oBrowse ALIAS 'INETUPS'                 ;
		SIZE 130,18                                     ;
		FONT '10.COURIER NEW'  ;
		PRESENTATION aPres ;
		WHEN {|| IIF( !EMPTY(oRecord:salesman),.t. ,.f. )}  ;
		DATALINK{|| _loadinetup(@oRecord),lSearchServ:=.f.,;
		DC_GETREFRESH(GETLIST)}

	DCBROWSECOL FIELD INETUPS->DATERECVD                     ;
		WIDTH 15 HEADER "Date Recvd" PARENT oBrowse


	DCBROWSECOL FIELD INETUPS->FIRSTNAME                     ;
		WIDTH 15 HEADER "First Name" PARENT oBrowse

	DCBROWSECOL FIELD INETUPS->LASTNAME                     ;
		WIDTH 15 HEADER "Last Name" PARENT oBrowse

	DCBROWSECOL FIELD INETUPS->Email                     ;
		WIDTH 15 HEADER "eMail" PARENT oBrowse

	DCBROWSECOL FIELD INETUPS->city                     ;
		WIDTH 15 HEADER "City" PARENT oBrowse
	DCBROWSECOL FIELD INETUPS->state                     ;
		WIDTH 3 HEADER "State" PARENT oBrowse
	DCBROWSECOL FIELD INETUPS->SM                     ;
		WIDTH 3 HEADER "SP_ID" PARENT oBrowse
	DCBROWSECOL FIELD INETUPS->CONLAST                     ;
		WIDTH 5 HEADER "SP_Name" PARENT oBrowse

	DCBROWSECOL FIELD INETUPS->OPTOUT                     ;
		WIDTH 5 HEADER "Status" PARENT oBrowse



  @  30,5 DCGET cUpMessage EDITPROTECT{|| TRUE }  GETCOLOR 1,BD_LEMONCREME

  @ 32,5 DCSAY 'Press OK to Continue or Cancel to Exit'



	dcread gui APPWINDOW MAINWINDOW():DRAWINGAREA ADDBUTTONS to xstatus fit options getoptions title 'Enter New Up '
	IF !xStatus
      AboCloseAll()
		return nil
	ENDIF

	IF empty(oRecord:salesman)
		@ 2,5 DCSAY 'YOU MUST ENTER YOUR SALES PERSON ID'
		@ 4,5 DCSAY '                  Enter Valid Salesmans Initials' Get oRecord:SALESMAN Valid{|| Chksm(oRecord:SALESMAN)} ;
		getcolor{|| IIF(EMPTY(oRecord:salesman),{1,BD_SALMON} ,{1,BD_KEYLIME}  ) }   Picture '@!'
		dcread gui MODAL ADDBUTTONS to xstatus fit options getoptions title 'Enter New Up ' EVAL{|o| setappwindow(o)}
	   IF !xStatus
       AboCloseAll()
		 return nil
	   ENDIF
	ENDIF

	IF empty(oRecord:UPID)
		@ 2,5 DCSAY 'YOU MUST ENTER AN ID FOR THIS CUSTOMER- 7 DIGIT PHONE#  OR PARTIAL EMAIL '
		@ 4,5 DCSAY '                  Enter UP ID' Get oRecord:UPID Valid{|| IIF( EMPTY(oRecord:UPID),.F. ,.T. )}
		dcread gui MODAL ADDBUTTONS to xstatus fit options getoptions title 'Enter New Up ' EVAL{|o| setappwindow(o)}
	   IF !xStatus
       AboCloseAll()
		 return nil
	   ENDIF
	ENDIF

	IF empty(oRecord:LASTNAME)
		@ 2,5 DCSAY 'YOU MUST ENTER A LAST NAME THIS CUSTOMER- '
		@ 4,5 DCSAY '                  Enter Last Name' Get oRecord:lastname Valid{|| IIF( EMPTY(oRecord:lastname),.F. ,.T. )}
		dcread gui MODAL ADDBUTTONS to xstatus fit options getoptions title 'Enter New Up ' EVAL{|o| setappwindow(o)}
	   IF !xStatus
       AboCloseAll()
		 return nil
	   ENDIF
	ENDIF






	NEWUPS->(DC_DBGATHER(oRecord,.T.))
	nRec:=NEWUPS->(RECNO())
	// destroy record object after creating
	 oRecord:destroy
	/// if up is found- go to Mainup()
	IF lBeback
		upquote(nRec)
		RETURN NIL
	ENDIF
	/*
	NEWUP(@xStatus,@oRecord)
		  IF !xStatus
			RETURN NIL
		  ENDIF */

	///now run main up entry edit screen . a new record object will be created
	upquote(nRec)
	//select newups
   AboCloseAll()
RETURN .T.

STATIC FUNCTION _searchinetups(oRecord)
	local getlist:={}, cLastName:=space(10)
	local getoptions
	@ 10,0 DCSAY 'Enter the name to search for' get cLastName PICTURE '@!'
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
	DCREAD GUI MODAL ADDBUTTONS ENTEREXIT OPTIONS GETOPTIONS FIT TITLE 'Search iNet' EVAL{|o| SETAPPWINDOW(o)}
	_chkinetups(cLastName,@oRecord,.T.)

	return nil

STATIC function _lookupup(oRecord,lSearchserv,cUpmessage,lBeback)
	local clacon:='Y'
	local getoptions
	local cSrchnm:=space(20), xStatus
	LOCAL GETLIST:={}
	LOCAL nREC:=0
	local cHomePhone:=' '
	DCGEToptions autoresize saywidth 0 sayfont '12.Courier New' getfont '12.Courier New'
	///  accept blank phn number  assuming it will be populated by inet ups or service cust search
	IF EMPTY(oRecord:UPID)
	  	//dc_winalert('You must enter a 7 Digit Phone number here.!')
	  	RETURN .t.
	ELSE   /// look up phone number
			SELECT NEWUPS
			ORDSETFOCUS('NEWUPIN')
			SEEK oRecord:UPID
   		IF FOUND()
				IF(Date()-(NEWUPS->DATEIN)) < 91          ///// ASSUME GO TO MAIN UP FOR BEBACK
			   	lBeBack=.t.
   			ENDIF
				/// ASK IF THIS IS A BEBACK OR A NEW UP RECORD
				@ 5,0 DCSAY 'The system has found an existing Up record in the current file. If this is a BeBack entry'
				@ 6,0 DCSAY 'then simply enter a Y below and you will be SET UP TO BE taken to the Maintain Up screen.'
				@ 7,0 DCSAY 'If this is NOT a Be Back but is a new sales opportunity or perhaps a second purchase'
				@ 8,0 DCSAY 'by this customer then enter N and you will be set up to record a NEW initial up record.'
				@ 10,0 DCCHECKBOX lBeback PROMPT 'UnCheck this box if this is NOT a Be Back'
				  DCREAD GUI MODAL ADDBUTTONS ENTEREXIT OPTIONS GETOPTIONS FIT TITLE 'Up Record Found' EVAL{|o| SETAPPWINDOW(o)}
				IF lBeBack
					cUpMessage:='This Up will be recorded as a Be Back'
					RETURN TRUE
				ENDIF
				oRecord:AREA=NEWUPS->AREA
				oRecord:FIRSTNAME:=NEWUPS->FIRSTNAME
				oRecord:LASTNAME:=NEWUPS->LASTNAME
				oRecord:STREET:=NEWUPS->STREET
				oRecord:CITYSTATE:=NEWUPS->CITYSTATE
				oRecord:STATE:=NEWUPS->STATE
				oRecord:ZIPCODE:=NEWUPS->ZIPCODE
				oRecord:EMAIL:=NEWUPS->EMAIL
				oRecord:WORKPHN:=NEWUPS->WORKPHN
				oRecord:CELLPHONE:=NEWUPS->cellphone
				cUpMessage:='This Up will be recorded as a New Up Opportunity'

			   RETURN .t.
   		ENDIF
		/// now search hnewups
   		SELECT HNEWUPS
   		SEEK oRecord:UPID
   		IF FOUND()
			 oRecord:AREA=HNEWUPS->AREA
				oRecord:FIRSTNAME:=HNEWUPS->FIRSTNAME
				oRecord:LASTNAME:=HNEWUPS->LASTNAME
				oRecord:STREET:=HNEWUPS->STREET
				oRecord:CITYSTATE:=HNEWUPS->CITYSTATE
				IF HNEWUPS->(ISFIELDVAR('STATE'))
				 oRecord:STATE:=HNEWUPS->STATE
				ENDIF
				oRecord:ZIPCODE:=HNEWUPS->ZIPCODE
				oRecord:EMAIL:=HNEWUPS->EMAIL
				oRecord:WORKPHN:=HNEWUPS->WORKPHN
				oRecord:CELLPHONE:=HNEWUPS->cellphone
			  cUpMessage:='This Up will be recorded as a New Up Opportunity'
   		RETURN .t.
  		ENDIF

  		IF !FOUND()
			IF lSearchserv
					@ 10,0 DCSAY 'Search Service Customer Base Y/N' get clacon PICTURE '@! L'
					dcread gui MODAL enterexit to xstatus options getoptions fit title 'Search Customer Base'   ;
						EVAl{|o| setappwindow(o)}
					if !xstatus
						return xStatus
					endif
					IF clACON='N'
						RETURN .T.
					ENDIF
					IF clACON = "Y"
						@ 3,0 DCSAY 'Enter Search Name Up To 20 Characters Esc=exit'  GET cSRCHNM PICTURE "!!!!!!!!!!!!!!!!!!!!"   ;
							GETTOOLTIP('YOU MAY ENTER UP TO 20 CHARACTERS, HOWEVER 3 OR 4 IS USUALLY ENOUGH ;TO PLACE THE FILE POINTER AT THE APPROXIMATE POINT TO EASILY ISOLATE THE RECORD NEEDED.;  YOU MAY PRESS ESCAPE TO EXIT ')
						DCREAD GUI modal TO XSTATUS BUTTONS 2 ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'SEARCH FOR NAME  ' ;
							eval{|o| setappwindow(o)}
						IF !XSTATUS
							RETURN xStatus
						ENDIF
						SELECT SERVCUST
						SET SOFTSEEK ON
						LOOKSERV(@nREC,cSRCHNM)
						SET SOFTSEEK OFF
					ENDIF  &    &ACON
					IF nREC <> 0
						SERVCUST->(DBGOTO(nREC))                                  && IF ONE IS SELECTED
						oRecord:LASTNAME:=SERVCUST->LASTNAME
						oRecord:FIRSTNAME:=SERVCUST->FIRSTNAME
						oRecord:STREET:=SERVCUST->STREET
						oRecord:CITYSTATE:=SERVCUST->CITYSTATE
						oRecord:STATE:=SERVCUST->STATE
						oRecord:ZIPCODE:=SERVCUST->ZIP
						oRecord:WORKPHN:=STR(SERVCUST->WPHONE)
						oRecord:EMAIL:=SERVCUST->EMAIL
						chomePhone:=alltrim(str(SERVCUST->phone))          /// convert numeric to string and see if there is an area code
						IF Len(alltrim(chomePhone)) > 7
							oRecord:AREA:=substr(alltrim(chomePhone),1,3)
							oRecord:UPID:=substr(alltrim(chomePhone),4,7)
						Endif
						IF len(cHomePhone)=7
							oRecord:UPID:=cHomePhone
						ENDIF

						oRecord:CELLPHONE:=SERVCUST->cell
					ENDIF
			endif /// lsearchserv
		endif
      cUpMessage:='This Up will be recorded as a New Up Opportunity'
ENDIF   /// not empty
RETURN .T.


STATIC FUNCTION UPSEARCHBUTTON(oRecord,lSearchserv)
local getlist:={}, xStatus  , cHomephone
local getoptions
local nRec:=0
local cSRCHNM:=space(20)
DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
@ 3,0 DCSAY 'Enter Search Name Up To 20 Characters Esc=exit'  GET cSRCHNM PICTURE "!!!!!!!!!!!!!!!!!!!!"   ;
			GETTOOLTIP('YOU MAY ENTER UP TO 20 CHARACTERS, HOWEVER 3 OR 4 IS USUALLY ENOUGH ;TO PLACE THE FILE POINTER AT THE APPROXIMATE POINT TO EASILY ISOLATE THE RECORD NEEDED.;  YOU MAY PRESS ESCAPE TO EXIT ')
DCREAD GUI modal TO XSTATUS BUTTONS 2 ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'SEARCH FOR NAME  ' ;
			eval{|o| setappwindow(o)}
IF !XSTATUS
	RETURN xStatus
ENDIF
lSearchserv:=.f.
SELECT SERVCUST
SET SOFTSEEK ON
	LOOKSERV(@nREC,cSRCHNM)
	SET SOFTSEEK OFF
	IF nREC <> 0
	 GOTO nREC                                  && IF ONE IS SELECTED
	 oRecord:LASTNAME:=SERVCUST->LASTNAME
	 oRecord:FIRSTNAME:=SERVCUST->FIRSTNAME
	 oRecord:STREET:=SERVCUST->STREET
	 oRecord:CITYSTATE:=SERVCUST->CITYSTATE
	 oRecord:STATE:=SERVCUST->STATE
	 oRecord:ZIPCODE:=SERVCUST->ZIP
	 oRecord:WORKPHN:=SERVCUST->WPHONEALPH
	 oRecord:eMAIL:=SERVCUST->EMAIL
	 chomePhone:=SERVCUST->HphoneALPH
	 IF Len(alltrim(chomePhone)) > 7
		oRecord:AREA:=substr(alltrim(chomePhone),1,3)
		oRecord:UPID:=substr(alltrim(chomePhone),4,7)
	 Endif
	 IF Len(alltrim(chomePhone)) = 7
		oRecord:UPID:=alltrim(chomePhone)
	 Endif
	oRecord:CELLPHONE:=SERVCUST->cell
  ENDIF
RETURN .T.






FUNCTION MUSTDEMO(aUpChars)
	IF NEWUPS->NEWUSED='N'
		RETURN .T.
	ENDIF
	IF EMPTY(aUpChars[UPSC_CDEMODSTKNO])
		DC_winALERT('USED VEHICLE DEMOS MUST RECORD STOCK NUMBER SHOWN')
		RETURN .F.
	ENDIF
RETURN .T.


// END UPENTRY
FUNCTION NEWUP(xstatus,oRecord)
	local getlist:={}  , botoolbar ,nRec:=0  ,oUpInfo,oIntinfo,oOtherinfo
	local getoptions
	local cUpstat:='A'
	local dDateToday :=date()
	local dDatesold:=ctod('  /  /  ')
	DCGEToptions saywidth 0 sayfont '12.Courier New' getfont '12.Courier New' autoresize NOTABSTOP

	/// SET DEFAULTS
	oRecord:UPSTATUS:='A'
	oRecord:ORIGSOURCE:='F'

	@ 2,0 DCGROUP oUpInfo CAPTION 'Prospect Demographics' size 140,11
	@ 1,5 DCSAY ' PHONE#/eMail'  get oRecord:UPID VALID{|| IIF(EMPTY(oRecord:UPID),.F. ,.T. )} PARENT oUpInfo
	@ 1,50 DCSAY 'Enter Area Code' GET oRecord:AREA   VALID {|| SETAPPFOCUS(DC_GETOBJECT(GETLIST,'LASTNAME')),.T.};
		PARENT oUpInfo

	@ 2,5 DCSAY '     Lastname' GET oRecord:LASTNAME SAYCOLOR 9,0   PICTURE '@!' PARENT oUpInfo GETID 'LASTNAME';
		  VALID{|| _CHKINETUPS(oRecord:lastname,oRecord,.f.),DC_GETREFRESH(GETLIST),.T.} ;   /// chk to see if this is an internet up
		  getcolor{|| IIF( empty(oRecord:lastname),{1,BD_SALMON} ,{1,BD_KEYLIME} )}
	@ 3,5 DCSAY '    Firstname' GET oRecord:FIRSTNAME  SAYCOLOR 9,0  PICTURE '@!' PARENT oUpInfo

	@ 3,70 DCSAY 'Zipcode' GET oRecord:ZIPCODE SAYCOLOR 9,0;
		getcolor{|| IIF( empty(oRecord:zipcode),{1,BD_SALMON} ,{1,BD_KEYLIME} )};
		VALID{|| SEEKZIP(@oRecord),DC_GETREFRESH(GETLIST),.T.};
		POPUP{|| BROWGETZIP(@oRecord),DC_GETREFRESH(GETLIST)};
		POPCAPTION 'ZipSearch ' POPWIDTH 100 POPFONT '10.Courier New' POPSTYLE 1 PARENT oUpInfo
	@ 4,5 DCSAY '       Street' GET oRecord:STREET SAYCOLOR 9,0    PICTURE '@!'  PARENT oUpInfo
	@ 5,5 DCSAY '         City' GET oRecord:CITYSTATE  SAYCOLOR 9,0 PICTURE '@!' PARENT oUpInfo
	@ 5,50 DCSAY 'State' GET oRecord:STATE   PARENT oUpInfo
	@ 6,5 DCSAY ' Last Date In'     PARENT oUpInfo SAYCOLOR 9,0
	@ 6,30 DCSAY oRecord:beback3  PARENT oUpInfo

	IF (dDATetoday-oRecord:DATEIN) < 1000
		@ 7,50 DCSAY 'Days since here last  '+ STR((dDATetoday-(oRecord:DATEIN)) )     PARENT oUpInfo
	ENDIF
	@ 8,5 DCSAY '        eMail' GET oRecord:EMAIL SAYCOLOR 2,0 PICTURE '@!' valid{|| _chkemailaddress(@oRecord,.F.)}   PARENT oUpInfo
	@ 9,5 DCSAY '   Work Phone' GET oRecord:WORKPHN  SAYCOLOR 2,0  PICTURE '@R 999-999-9999-AAAA'  PARENT oUpInfo
	@ 10,5 DCSAY '   Cell Phone' GET oRecord:CELLPHONE  PICTURE '@R 999-999-9999'  PARENT oUpInfo  VALID {|| SETAPPFOCUS(DC_GETOBJECT(GETLIST,'VEHICLE')),.T.}


	@ 13,0 DCGROUP oIntInfo Caption 'Visit Information' size 140,8
	@ 1,5 DCSAY '                                                      Vehicle Interest' GET oRecord:INTEREST PARENT oIntInfo GETID 'VEHICLE';
		SAYCOLOR 9,0;
		VALID{||CHKINTR(@oRecord:INTEREST,@oRecord:FRANCHISE),dc_getrefresh(getlist),.T.}

	@ 2,5 DCSAY '             Enter Source (F)loor Traffic  (I)nternet Lead  (P)hone Up' get oRecord:ORIGSOURCE VALID{|| IIF( oRecord:ORIGSOURCE $('FIP'),.T. ,.F. )}  ;
		PICTURE '@!' PARENT oIntInfo SAYCOLOR 9,0

	@ 3,5 DCSAY 'For Ups that have PHYSICALLY been in the store use Up Type 123 or 4'  SAYCOLOR 1,BD_LEMONCREME PARENT oIntInfo
	@ 4,5 DCSAY 'For Ups that have NOT been in the store yet use Up Type 5 or 6     '  SAYCOLOR 1,BD_LEMONCREME PARENT oIntInfo
	@ 5,5 DCSAY 'UP Type 1=NewCust 2=ExistingCust 3=Beback/1 4=Beback/2. 5=iNet 6=Phone' GET oRecord:UPTYPE PICTURE '9' PARENT oIntInfo;
		 RANGE 1,6  SAYCOLOR 9,0
	@ 6,5 DCSAY '                        Enter N For New Vehicle Or U For Used Vehicle.' GET oRecord:NEWUSED PICTURE '@!';
		 VALID{|| CHKNU(oRecord:NEWUSED)} PICTURE '@!' PARENT oIntInfo SAYCOLOR 9,0    ;
		 getcolor{|| IIF( empty(oRecord:NEWUSED),{1,BD_SALMON} ,{1,BD_KEYLIME} )}


	@ 7,5 DCSAY '                   Enter Advertising Source or Press Enter for Choices' GET oRecord:ADSOURCE ;
		SAYCOLOR 9,0;
		PICTURE '@!' PARENT oIntInfo ;
		valid {|| aslk(@oRecord:ADSOURCE),dc_getrefresh(getlist), SETAPPFOCUS(DC_GETOBJECT(GETLIST,'COMMENTS')),.T.}

  @ 23,0 DCGROUP oOtherInfo CAPTION 'Comments and Other Info' size 140,7.5
	@ 1,5 DCSAY 'Enter Comments' GET oRecord:COMMENTS  SAYCOLOR 9,0 PARENT oOtherInfo GETID 'COMMENTS'
	@ 2,5 DCSAY 'More Comments ' GET oRecord:COMMENT2  SAYCOLOR 9,0   PARENT oOtherInfo

	@ 3,5 DCSAY 'T.O.Manager   '  GET oRecord:TOMGR PICTURE "!!!"  VALID{|| IIF(!EMPTY(oRecord:TOMGR),oRecord:TOED:=.T.,oRecord:TOED:=.F.),;
		DC_GETREFRESH(GETLIST),.T.} SAYCOLOR 9,0 PARENT oOtherInfo

	@ 4,5 DCSAY 'Demo Ride Y/N ' GET oRecord:DEMO PICTURE '@! L'  VALID{|| IIF(oRecord:DEMO='Y',oRecord:DEMOED:=.T.,oRecord:DEMOED:=.F.),.T.};
		 SAYCOLOR 9,0 PARENT oOtherInfo

	@ 5,5 DCSAY 'Enter Follow up Date             ' GET oRecord:TICKLE PARENT oOtherInfo

	@ 6,5 DCSAY 'Enter the status of this up.  (A)ctive (R)escue (F)uture (D)ead  (S)old '  get oRecord:UPSTATUS PICTURE '@!' PARENT oOtherInfo ;
		valid{|| editupstat(@oRecord,@Getlist)};
		sayfont '12.Courier New' ;
		saycolor 1,BD_LEMONCREME



	DCREAD GUI APPWINDOW MAINWINDOW():DRAWINGAREA TO XSTATUS ADDBUTTONS FIT OPTIONS GETOPTIONS TITLE 'Prospect Entry Info'
	/// IF ESCAPE IS PRESSED RETURN NIL.
	 IF !XSTATUS
		RETURN NIL
	ENDIF

	NEWUPS->(DC_DBGATHER(oRecord,.t.))


RETURN NIL

STATIC FUNCTION _chkemailaddress(oRecord,ALLOWOVERRIDE)
	local nEmail:=1
	DEFAULT ALLOWOVERRIDE:=.F.
	IF EMPTY(oRecord:EMAIL)
		dc_winalert('Please enter a valid Email address containing a @ and a . or WNG or DNH .')
		RETURN FALSE
	ENDIF

	IF UPPER(ALLTRIM(oRecord:EMAIL))='DNH'
		oRecord:EMAIL='DNH@A.COM'
		return TRUE
	endif
	IF UPPER(ALLTRIM(oRecord:EMAIL))='WNG'
		oRecord:EMAIL='WNG@A.COM'
		return TRUE
	endif

	if !fexists('getemail.txt')                  /// only force a look if getemail .txt is present. return t if it is not
		return .t.
	endif
	IF ALLOWOVERRIDE=.T.
		return .t.
	endif

	IF !EMPTY(oRecord:EMAIL)
	 SELECT INETUPS
	 ORDSETFOCUS('INETUPS')
	 SEEK oRecord:EMAIL
	 IF FOUND()
		oRecord:ORIGSOURCE='I'
		RETURN TRUE
	 ENDIF
	ENDIF


	nEmail:=AT('@',oRecord:EMAIL,2)              /// make sure address is valid
	if nEmail=0
		dc_winalert('Please enter a valid Email address containing a @ and a . or WNG or DNH .')
		return .f.
	endif
	nEmail:=0
	nEmail:=AT('.',oRecord:eMAIL,2)              /// make sure address is valid
	if nEmail=0
		dc_winalert('Please enter a valid Email address containing a @ and a . or WNG or DNH .')
		return .f.
	endif


return .t.


static Function _CHKINETUPS(cLastName,oRecord,lSetsofton)
	local getlist:={},oBrowse, apres , aScopearray:={} , getoptions , xStatus
	DEFAULT lSetsofton:=.f.
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE

	SELECT INETUPS
	ORDSETFOCUS('INETLAST')

	IF !EMPTY(oRecord:ORIGSOURCE)
		RETURN TRUE
	ENDIF

	IF !lSetSofton
	 INETUPS->(DBGOTOP())
	 SEEK upper(cLastName)
	 IF !FOUND()
		RETURN TRUE
	 ENDIF
	ENDIF

	IF lSetSofton
		SET SOFTSEEK ON
		seek cLastname
		SET SOFTSEEK OFF
	ENDIF


		 aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_ROWHEIGHT, -1 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 15 }  }              /* Cell Height */

	@ 1,1 DCSAY 'You May Double click on a selected record to access the data to create an up record.' SAYCOLOR 1,5
	@ 3,1 DCSAY 'If NO record is to be selected press Esc Key or the Cancel Button'  SAYCOLOR 1,7


	@ 5,1 DCBROWSE oBrowse ALIAS 'INETUPS'                 ;
		SIZE 110,20                                     ;
		FONT '10.COURIER NEW'  ;
		PRESENTATION aPres ;
		DATALINK{|| _loadinetup(@oRecord),;
		DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}

	DCBROWSECOL FIELD INETUPS->SM                     ;
		WIDTH 3 HEADER "SP_ID" PARENT oBrowse
	DCBROWSECOL FIELD INETUPS->CONLAST                     ;
		WIDTH 5 HEADER "SP_Name" PARENT oBrowse

	DCBROWSECOL FIELD INETUPS->FIRSTNAME                     ;
		WIDTH 15 HEADER "First Name" PARENT oBrowse

	DCBROWSECOL FIELD INETUPS->LASTNAME                     ;
		WIDTH 15 HEADER "Last Name" PARENT oBrowse

	DCBROWSECOL FIELD INETUPS->Email                     ;
		WIDTH 15 HEADER "eMail" PARENT oBrowse

	DCBROWSECOL FIELD INETUPS->city                     ;
		WIDTH 15 HEADER "City" PARENT oBrowse
	DCBROWSECOL FIELD INETUPS->state                     ;
		WIDTH 3 HEADER "State" PARENT oBrowse

	DCREAD GUI MODAL;
		TO XSTATUS;
		EVAL {|o| setappwindow(o)};
		FIT TITLE 'iNet Up Browser ' ;
		OPTIONS GETOPTIONS;
		ADDBUTTONS

  IF !xStatus
   return nil
  ENDIF

  _loadinetup(@oRecord)
  DC_CLRSCOPE()
RETURN NIL

static function _loadinetup(oRecord)
	LOCAL cUpid:='',cArea:='', cPhone:='',nRec:=0
	SELECT INETUPS
	oRecord:FIRSTNAME:=INETUPS->FIRSTNAME
	oRecord:LASTNAME:=INETUPS->LASTNAME
	oRecord:STREET:=INETUPS->STREET
	oRecord:CITYSTATE:=INETUPS->CITY
	oRecord:ZIPCODE:=INETUPS->ZIPCODE
	IF EMPTY(oRecord:salesman)
	 oRecord:SALESMAN:=INETUPS->SM
	ENDIF
	oRecord:COMMENTS:='InterNet Up'
	oRecord:ADSOURCE:=INETUPS->SOURCE
	oRecord:EMAIL:=INETUPS->EMAIL
	oRecord:UPTYPE:=1
	//aUpChars[UPSC_CWORKPHONE]:=INETUPS->DAYPHONE
	oRecord:ORIGSOURCE:='I'
	IF !EMPTY(INETUPS->DAYPHONE)
		cPhone:=strtran(INETUPS->DAYPHONE,'-','')
		cPhone:=strtran(cPhone,'(','')
		cPhone:=strtran(cPhone,')','')
		cPhone:=alltrim(cPhone)
		cArea:=substr(cPhone,1,3)
		cUpId:=substr(cPhone,4,7)
		oRecord:UPID:=cUpid
		oRecord:AREA:=cArea
	  ELSE
		cPhone:=strtran(INETUPS->EVEPHON,'-','')
		cPhone:=strtran(cPhone,'(','')
		cPhone:=strtran(cPhone,')','')
		cPhone:=alltrim(cPhone)
		cArea:=substr(cPhone,1,3)
		cUpId:=substr(cPhone,4,7)
		oRecord:UPID:=cUpid
		oRecord:AREA:=cArea
	ENDIF
	SELECT INETUPS
	IF INETUPS->(DC_RECLOCK())
		nRec:=INETUPS->(RECNO())
		REPLACE INETUPS->OPTOUT WITH 'CONVERTED'
		REPLACE INETUPS->SM WITH oRecord:salesman
		INETUPS->(DBRUNLOCK(nRec))
	ENDIF

RETURN NIL





FUNCTION UPNOTES()
	SELECT NEWUPS
	IF NEWUPS->(DC_RECLOCK())
		REPLACE NEWUPS->NOTES WITH DC_MEMOEDIT(NEWUPS->NOTES,1,10,20,110,.T.)
		NEWUPS->(DBCOMMIT())
		NEWUPS->(DBRUNLOCK())
	ENDIF
RETURN .T.

FUNCTION CHKINTR(cVehInterest,cFranchise)
	IF empty(cVehInterest)
		RETURN FALSE
	ENDIF
	SELECT CLASSORT
	SEEK cVehInterest
	IF !FOUND()
		SELECT CLASSORT
		CLASSORT->(DBGOTOP())
		nprodlk(@cVehInterest,@cFranchise)
		RETURN .T.
	ENDIF
RETURN .T.


FUNCTION PRINTUPSHEET(nREC)
	local aATTR,bATTR,oPrinter

	SELECT NEWUPS
	GOTO nREC
	DCPRINT ON SIZE 66,132 USEDEFAULT TO oPrinter FONT '10.Courier New'
	IF VALTYPE(OPRINTER) # 'O' .OR. !OPRINTER:LACTIVE
				RETURN NIL
	ENDIF
	aATTR:=ARRAY(GRA_AL_COUNT)
	aAttr[GRA_AL_COLOR]:=GRA_CLR_BLACK
	battr:=array(GRA_AA_COUNT)
	bAttr[GRA_AA_BACKCOLOR]:=GRA_CLR_PALEGRAY



	@ 1,0 DCPRINT SAY 'SalesPerson ' + NEWUPS->salesman
	@ 3,20 DCPRINT SAY ALLTRIM(CONAME()) fONT '24.Courier New Bold'
	@ 5,5 DCPRINT SAY 'Date '+ DTOC(DATE()) font '12.Courier New Bold'
	@ 5,50 DCPRINT SAY 'UPTYPE(1,2,3,4,5=iNet)  ' + STR(NEWUPS->uptype)  fONT '12.Courier New'
	@ 7,5 DCPRINT SAY 'Name '+ NEWUPS->firstname+' '+NEWUPS->lastname   fONT '12.Courier New Bold'
	@ 8,5 DCPRINT SAY 'Address ' + NEWUPS->street    fONT '12.Courier New '
	@ 8,50 DCPRINT SAY 'Advertising Source '+ NEWUPS->ADSOURCE     fONT '12.Courier New'
	@ 9,5 DCPRINT SAY 'City, State, Zip '+ NEWUPS->Citystate+' '+NEWUPS->zipcode  fONT '12.Courier New'
	@ 10,5 DCPRINT SAY 'Daytime Phone '+ NEWUPS->area+' '+NEWUPS->UPID   fONT '12.Courier New'
	@ 11,5 DCPRINT SAY 'Work Phone '+NEWUPS->workphn fONT '12.Courier New'
	@ 11,50 DCPRINT SAY 'Cell Phone '+NEWUPS->cellphone fONT '12.Courier New'
	@ 12,5 DCPRINT SAY 'EMAIL '+ NEWUPS->email        fONT '12.Courier New'
	@ 13,5 DCPRINT SAY 'Preferred Contact Method ' +NEWUPS->PREFMETHOD
	@ 14,5 DCPRINT SAY 'Vehicle Information ' + NEWUPS->interest        fONT '12.Courier New'
	@ 15,5 DCPRINT SAY 'YR/TYPE '+ NEWUPS->yrquoted+' '+NEWUPS->vehquoted    FONT '12.Courier New'
	@ 16,5 DCPRINT SAY 'Price__'+str(NEWUPS->quoted)+'  Stock Number '+NEWUPS->DEMOStK +'  Color_________' FONT '12.Courier New'
	@ 17,5 DCPRINT SAY 'Trade Info Year Make Model____________________________________________________________'     FONT '12.Courier New'
	@ 18,5 DCPRINT SAY 'Comments '
	@ 18,20 DCPRINT SAY NEWUPS->comments
	@ 19,20 DCPRINT SAY NEWUPS->comment2
	@ 20,5 DCPRINT SAY 'Conventional Financing' FONT '10.Courier New Bold'
	@ 21,70 DCPRINT SAY 'Lease/Balloon Purchase' FONT '10.Courier New Bold'
	@ 22,5 DCPRINT SAY '          Downpayment'            FONT '10.Courier New Bold'
	@ 22,70 DCPRINT SAY '          Downpayment'  FONT '10.Courier New Bold'
	@ 23,5 DCPRINT SAY ' Term'  FONT '10.Courier New Bold'
	@ 23,16 DCPRINT SAY ' $500' FONT '10.Courier New Bold'
	@ 23,27 DCPRINT SAY ' $1000' FONT '10.Courier New Bold'
	@ 23,38 DCPRINT SAY ' $2000' FONT '10.Courier New Bold'
	@ 23,49 DCPRINT SAY ' $    ' FONT '10.Courier New Bold'

	@ 23,70 DCPRINT SAY ' Term'  FONT '10.Courier New Bold'
	@ 23,81 DCPRINT SAY ' $500' FONT '10.Courier New Bold'
	@ 23,92 DCPRINT SAY ' $1000' FONT '10.Courier New Bold'
	@ 23,103 DCPRINT SAY ' $2000' FONT '10.Courier New Bold'
	@ 23,114 DCPRINT SAY ' $    ' FONT '10.Courier New Bold'

	@ 25,5,27,15 dcprint box LINEATTR aATTR  AREAATTR bAttr
	@ 28,5,30,15 dcprint box LINEATTR aATTR  AREAATTR bAttr
	@ 31,5,33,15 dcprint box LINEATTR aATTR  AREAATTR bAttr

	@ 25,16,27,26 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 28,16,30,26 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 31,16,33,26 dcprint box LINEATTR aATTR AREAATTR bAttr

	@ 25,27,27,37 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 28,27,30,37 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 31,27,33,37 dcprint box LINEATTR aATTR AREAATTR bAttr

	@ 25,38,27,48 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 28,38,30,48 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 31,38,33,48 dcprint box LINEATTR aATTR AREAATTR bAttr

	@ 25,49,27,59 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 28,49,30,59 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 31,49,33,59 dcprint box LINEATTR aATTR AREAATTR bAttr

	@ 25,70,27,80 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 28,70,30,80 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 31,70,33,80 dcprint box LINEATTR aATTR AREAATTR bAttr


	@ 25,81,27,91 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 28,81,30,91 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 31,81,33,91 dcprint box LINEATTR aATTR AREAATTR bAttr

	@ 25,92,27,102 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 28,92,30,102 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 31,92,33,102 dcprint box LINEATTR aATTR AREAATTR bAttr

	@ 25,103,27,113 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 28,103,30,113 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 31,103,33,113 dcprint box LINEATTR aATTR AREAATTR bAttr

	@ 25,114,27,124 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 28,114,30,124 dcprint box LINEATTR aATTR AREAATTR bAttr
	@ 31,114,33,124 dcprint box LINEATTR aATTR AREAATTR bAttr


	@ 64,0 DCPRINT SAY 'Ad Source' FONT '8.Courier New Bold'
	select adsource
	adsource->(dbgotop())
	do while !adsource->(eof())
		@ dc_printerrow(),dc_printercol()+4 DCPRINT SAY alltrim(adsource->desc) FONT '8.Courier New Bold'
		IF DC_PRINTERCOL() > 124
			@ dc_printerrow()+1,0 DCPRINT SAY 'Ad Source' FONT '8.Courier New Bold'
		ENDIF
		skip alias adsource
	enddo

	dcprint off
return nil


FUNCTION FOLLOWUP()
	LOCAL GETLIST:={} , getoptions , xStatus
	LOCAL cSM:='  '
	local cProgram:='FOLLOWUP' , dDatex,dDatex2 , lPrevue , oPrinter
	local nCopies:=1,nOrient:=2,nDefmode:=4,cOutfile:='' , nPageno:=0
	TOP_MAR(3)
	BOT_MAR(2)
	PAGE_LEN(62)


	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.COURIER NEW' GETFONT '10.COURIER NEW' AUTORESIZE

	IF.NOT.SECURITY(@cProgram)
	 DC_winALERT('SECURITY VIOLATION')
    RETURN NIL
   ENDIF

dDatex:=dDatex2:=CTOD("  /  /  ")

IF !UseDb({'NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF

NEWUPS->(ORDSETFOCUS->('NEWUPFDT'))

DCGETOPTIONS AUTORESIZE SAYWIDTH 0
XSTATUS:=.T.
@ 10,0 DCSAY 'Enter Salesperson Code-Leave Blank For All ' Get cSM
@ 11,0 DCSAY 'Enter Date To Start Print For' Get dDatex
@ 12,0 DCSAY 'Enter Date To End Print For  ' GET dDatex2
DCREAD GUI TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA FIT ENTEREXIT OPTIONS GETOPTIONS TITLE 'DAILY FOLLOWUP REPORT'
IF !XSTATUS
 RETURN NIL
ENDIF
SELECT NEWUPS
SET SOFTSEEK ON
SEEK dDatex
IF !FOUND()
	DC_winALERT("NO UPS FOUND FOR THIS DATE")
   RETURN NIL
ENDIF
IF !EMPTY(cSM)
	SET FILTER TO NEWUPS->SALESMAN=cSM
ENDIF

 IF !Print_Choice('FollowUp List', @nCopies,' ',.F.,@nOrient,'8.Courier New',@nDefMode,cOutfile )
		RETURN NIL
 ENDIF
 IF !PrintOn('FollowUp List', oPrinter, nOrient, '8.Courier New', nCopies)
		  RETURN NIL
 ENDIF


DO WHILE !NEWUPS->(EOF())
	IF NEWUPS->TICKLE >=dDatex
		IF NEWUPS->TICKLE <= dDatex2
			@ DC_PRINTERROW()+1,0 DCPRINT SAY NEWUPS->SALESMAN
			IF DCPAGEEJECT()
				@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Follow Up Report'
			ENDIF
			@ DC_PRINTERROW(),5 DCPRINT SAY NEWUPS->FIRSTNAME
			@ DC_PRINTERROW(),31 DCPRINT SAY NEWUPS->LASTNAME
			@ DC_PRINTERROW(),52 DCPRINT SAY NEWUPS->AREA
			@ DC_PRINTERROW(),63 DCPRINT SAY NEWUPS->UPID
			@ DC_PRINTERROW(),72 DCPRINT SAY NEWUPS->COMMENTS
			@ DC_PRINTERROW()+1,0 DCPRINT SAY '  '
			IF DCPAGEEJECT()
				@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Follow Up Report'
			ENDIF

			@ DC_PRINTERROW(),5 DCPRINT SAY NEWUPS->STREET
			@ DC_PRINTERROW(),31 DCPRINT SAY NEWUPS->UPTYPE
			@ DC_PRINTERROW(),33 DCPRINT SAY NEWUPS->INTEREST
			@ DC_PRINTERROW(),43 DCPRINT SAY NEWUPS->DATEIN
			@ DC_PRINTERROW(),53 DCPRINT SAY NEWUPS->TICKLE
			@ DC_PRINTERROW(),72 DCPRINT SAY NEWUPS->COMMENT2

			@ DC_PRINTERROW()+1,0 DCPRINT SAY '  '
			IF DCPAGEEJECT()
				@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Follow Up Report'
			ENDIF

			@ DC_PRINTERROW(),5 DCPRINT SAY NEWUPS->CITYSTATE
			@ DC_PRINTERROW(),36 DCPRINT SAY NEWUPS->ZIPCODE
			@ DC_PRINTERROW(),40 DCPRINT SAY NEWUPS->PREFMETHOD
			@ DC_PRINTERROW(),50 DCPRINT SAY NEWUPS->EMAIL
			@ DC_PRINTERROW()+1,0 DCPRINT SAY ''
			IF DCPAGEEJECT()
				@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Follow Up Report'
			ENDIF

		ENDIF
	ENDIF
	SKIP ALIAS NEWUPS
ENDDO
PRINTOFF(oPrinter)
CLOSE DATABASES
RETURN NIL

/// program for salesman daily planner
FUNCTION DailyPlan()
	LOCAL GETLIST:={} , getoptions , cProgram
	LOCAL cSM:=SPACE(2) , oDialog ,  xStatus
	local nCopies:=1,nOrient:=1,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
	local lFoundups:=.T. , dCutoff:=DATE()-30 ,lAnnivers:=.t.,lOneTwenty:=.t.,lDeliveries:=.t.,lExpires:=.t.,lInservice:=.t.,;
		lBirthday:=.t.
	MEMVAR nPage
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE SAYFONT '12.Courier New' GETFONT '12.Courier New'
	TOP_MAR(3)
	BOT_MAR(2)
	PAGE_LEN(62)
	nPage:=0

	IF !OPENUPSFILES()
		CLOSE ALL
		RETURN NIL
	ENDIF
	IF !UseDb({'SERVRESV'})
		 DC_WINALERT('Cannot Open Files     .')
	   	 RETURN NIL
	ENDIF
	/// SET FOCUS
	NEWUPS->(ORDSETFOCUS('NEWUPSM'))
	SERVRESV->(ORDSETFOCUS('SRVRESSM'))
	UPAPPTS->(ORDSETFOCUS('UPAPPTSM'))
	SERVCUST->(ORDSETFOCUS('SERVSM'))


	/// Get salesperson to print or allow empty for all
	@ 7,0 DCSAY 'Enter Salesmans code or -Leave Blank for all. ' GET cSM  PICTURE '!!'
	@ 9,0 DCSAY 'Enter Cutoff Date to go back for Active Ups   ' get dCutoff
	@ 11,0 DCCHECKBOX lOnetwenty PROMPT '120 Follow Up Calls and Appointments'
	@ 13,0 DCCHECKBOX lAnnivers PROMPT 'Delivery Anniveraries'
	@ 15,0 DCCHECKBOX lDeliveries PROMPT 'Current Deliveries'
	@ 17,0 DCCHECKBOX lInService PROMPT 'Customer In For Service/Incoming Appointments'
	@ 19,0 DCCHECKBOX lExpires PROMPT 'Finance Expirations'
	@ 21,0 DCCHECKBOX lBirthday PROMPT 'Buyers Actual Birthday'


	DCREAD GUI TO XSTATUS ADDBUTTONS APPWINDOW MAINWINDOW():DRAWINGAREA FIT ENTEREXIT OPTIONS GETOPTIONS TITLE 'DAILY FOLLOWUP REPORT'
	IF !XSTATUS
		RETURN NIL
	ENDIF
	// IF EMPTY DO A BATCH RUN FOR ALL
	if empty(cSM)
		IF !Print_Choice('Daily Planner', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	      RETURN NIL
      ENDIF
      IF !PrintOn('Daily Planner', oPrinter, nOrient, cFont, nCopies,.t.)
	      RETURN NIL
      ENDIF

		allplans(dCutoff,lAnnivers,lOneTwenty,lDeliveries,lExpires,lInservice,lBirthday,nDefmode)
		printoff(oPrinter)
		close all
		return nil
	endif


	// now process one plan only
	IF !Print_Choice('Daily Planner', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	      RETURN NIL
   ENDIF
   IF !PrintOn('Daily Planner', oPrinter, nOrient, cFont, nCopies,.t.)
	      RETURN NIL
   ENDIF

	oneplan(cSM,@lFoundups,dCutoff,lAnnivers,lOneTwenty,lDeliveries,lExpires,lInservice,lBirthday,nDefmode)
	if lFoundups
		printoff(oPrinter)
	endif
	close all
return nil

static function oneplan(cSM,lFoundups,dCutoff,lAnnivers,lOneTwenty,lDeliveries,lExpires,lInservice,lBirthday,nDefmode)
	local odialog ,cEmpID:=SPACE(3)
	local nPage:=0 , nXCoord:=0
	local xDate
	local lSome2Print:=.f.
	LOCAL aATTR:=ARRAY(GRA_AA_COUNT)
	LOCAL aLATTR:=ARRAY(GRA_AL_COUNT)
	LOCAL aNATTR:=ARRAY(GRA_AA_COUNT)
	TOP_MAR(3)
	BOT_MAR(2)
	PAGE_LEN(62)

	default lFoundups:=.t.

	aATTR[GRA_AA_COLOR]:=GRA_CLR_PALEGRAY
	aLATTR[GRA_AL_COLOR]:=GRA_CLR_BLACK
	aATTR[GRA_AA_SYMBOL]:=GRA_SYM_DIAG2

	select SM
	seek cSM
	IF found()
		cEmpID:=SM->ID
	ENDIF

	SELECT NEWUPS
	SEEK cSM
	IF !FOUND()
		lFoundups:=.f.
		//DC_winALERT("NO UPS FOUND FOR THIS SalesPerson"+' '+cSM)
		RETURN NIL
	ENDIF

   IF nDefmode > 2
	 odialog:=dc_waiton('Now Processing for '+cSM)
	ENDIF


	planhdr(@nPage)
	nXcoord:=DC_PRINTERROW()+1
	@ nXCoord,0,nXCoord+1.3,131 DCPRINT BOX LINEATTR ALATTR LINEWIDTH 1 FILL GRA_OUTLINEFILL AREAATTR AATTR
	@ DC_PRINTERROW()+1,49 DCPRINT SAY 'Active Prospects in Sales Process' font '8.Courier New Bold'

	/// START LOOP LOOKING FOR ACTIVE UPS
	/// PUT UP WINDOW DRESS IF PREVIEW IS not BEING USED

	DO WHILE NEWUPS->SALESMAN=cSM
		IF NEWUPS->DATEIN >=dCutoff
		 IF NEWUPS->upstatus $('AR')
			if dcpageeject()
				planhdr(@nPage)
			endif

			PrntPlan(TOP_MAR(),BOT_MAR(),PAGE_LEN(),@nPage,.f.) /// EACH LINE PRINT

			@ DC_PRINTERROW()+1,0 DCPRINT SAY ''
			if dcpageeject()
				planhdr(@nPage)
			endif
		 ENDIF
		ENDIF
		IF NEWUPS->UPStatus='F'
			IF NEWUPS->TICKLE <= DATE()
			  if dcpageeject()
				planhdr(@nPage)
		    	endif

			   PrntPlan(TOP_MAR(),BOT_MAR(),PAGE_LEN(),@nPAGE,.f.) /// EACH LINE PRINT

			   @ DC_PRINTERROW()+1,0 DCPRINT SAY ''
			   if dcpageeject()
				 planhdr(@nPage)
			   endif
		    ENDIF
		ENDIF                               /// upstatus
		SKIP ALIAS NEWUPS
	ENDDO

	SELECT NEWUPS
	SEEK cSM
	IF found()
		nXCoord:=DC_PRINTERROW()+1
		@nXCoord,0,nXCoord+1.3,131 DCPRINT BOX LINEATTR ALATTR LINEWIDTH 1 FILL GRA_OUTLINEFILL AREAATTR AATTR
		@ DC_PRINTERROW()+1,50 DCPRINT SAY 'Follow Up Calls for Tickler File' font '8.Courier New Bold'
		DO WHILE NEWUPS->SALESMAN=cSm
			IF NEWUPS->UPStatus='F'
			 IF NEWUPS->TICKLE <= DATE()
			  if dcpageeject()
			  	planhdr(@npage)
		     endif

			  PrntPlan(TOP_MAR(),BOT_MAR(),PAGE_LEN(),@nPAGE,.f.) /// EACH LINE PRINT

			  @ DC_PRINTERROW()+1,0 DCPRINT SAY ''
			  if dcpageeject()
				 planhdr(@npage)
			  endif
		    ENDIF
		   ENDIF                               /// upstatus
		   SKIP ALIAS NEWUPS
	   ENDDO
	ENDIF

	SELECT UPAPPTS
	SEEK cSM
	IF FOUND()
		nXCoord:=DC_PRINTERROW()+1
		@nXCoord,0,nXCoord+1.3,131 DCPRINT BOX LINEATTR ALATTR LINEWIDTH 1 FILL GRA_OUTLINEFILL AREAATTR AATTR
		@ DC_PRINTERROW()+1,50 DCPRINT SAY 'Appointments for Today '+ DTOC(DATE())+' '+CDOW(DATE()) font '8.Courier New Bold'
	ENDIF
	if dcpageeject()
				 planhdr(@npage)
	endif

	DO WHILE UPAPPTS->SM=cSm
		IF UPAPPTS->APPTDATE = DATE()
			@ DC_PRINTERROW()+1,1 DCPRINT SAY SUBSTR(UPAPPTS->LASTNAME,1,20)
			@ DC_PRINTERROW(),30 DCPRINT SAY SUBSTR(UPAPPTS->FIRSTNAME,1,15)
			@ DC_PRINTERROW(),50 DCPRINT SAY UPAPPTS->APPTTIME
			@ DC_PRINTERROW(),60 DCPRINT SAY UPAPPTS->PHONE
			@ DC_PRINTERROW(),75 DCPRINT SAY UPAPPTS->EMAIL
			IF dcpageeject()
				 planhdr(@npage)
			ENDIF
			@ DC_PRINTERROW()+1,1 DCPRINT SAY UPAPPTS->REASON
			SKIPALINE()
			IF dcpageeject()
				 planhdr(@npage)
			ENDIF
		ENDIF
		SKIP ALIAS UPAPPTS
	ENDDO

	IF lDeliveries
		select newups
		dbgotop()
		seek cSM
		IF found()
		 nXCoord:=DC_PRINTERROW()+1
		 @nXCoord,0,nXCoord+1.3,131 DCPRINT BOX LINEATTR ALATTR LINEWIDTH 1 FILL GRA_OUTLINEFILL AREAATTR AATTR
		 @ DC_PRINTERROW()+1,55 DCPRINT SAY 'Deliveries in Process' font '8.Courier New Bold'
			if dcpageeject()
				planhdr(@npage)
			endif
			DO WHILE NEWUPS->SALESMAN=cSM
				IF NEWUPS->DATEIN >=dCutoff
				 IF NEWUPS->upstatus ='S'
					if NEWUPS->billstat='I'
						if dcpageeject()
							planhdr(@npage)
						endif
						PrntPlan(TOP_MAR(),BOT_MAR(),PAGE_LEN(),@nPAGE,.t.) /// EACH LINE PRINT
						@ DC_PRINTERROW()+1,0 DCPRINT SAY ''
						if dcpageeject()
							planhdr(@npage)
						endif
					endif                                  /// billstat
				ENDIF    /// UPSTATUS
			  ENDIF	   /// > CUTOFF
			  SKIP ALIAS NEWUPS
			ENDDO

		  ENDIF
		  select newups
		  dbgotop()
		  seek cSM
		  IF found()
			SKIPALINE()
			if dcpageeject()
				planhdr(@npage)
			endif
			SKIPALINE()
			if dcpageeject()
				planhdr(@npage)
			endif
			nXCoord:=DC_PRINTERROW()+1
		   @nXCoord,0,nXCoord+1.3,131 DCPRINT BOX LINEATTR ALATTR LINEWIDTH 1 FILL GRA_OUTLINEFILL AREAATTR AATTR
		   @ DC_PRINTERROW()+1,51 DCPRINT SAY 'Follow Up After Delivery Calls' font '8.Courier New Bold'
			if dcpageeject()
				planhdr(@npage)
			endif
			DO WHILE NEWUPS->SALESMAN=cSM
				IF NEWUPS->DATEIN >=dCutoff
			    IF NEWUPS->upstatus ='S'
					if NEWUPS->billstat='D'
						IF empty(NEWUPS->SENDLTR)
						  if dcpageeject()
						   	planhdr(@npage)
						  endif
						  PrntPlan(TOP_MAR(),BOT_MAR(),PAGE_LEN(),@nPAGE,.t.) /// EACH LINE PRINT
						  @ DC_PRINTERROW()+1,0 DCPRINT SAY ''
						  if dcpageeject()
						   	planhdr(@npage)
						  endif
						ENDIF   /// SENDLTR
					endif                                  /// billstat
				ENDIF                                     /// upstatus =S
			  ENDIF  /// CUTOFF
			  SKIP ALIAS NEWUPS
			ENDDO
			dcprint eject
		  ENDIF

	ENDIF
	/// now look for 120 day calls
	IF lOnetwenty
		xdate:=Date()-120

		/// SEE IF THERE ARE ANY 120 DAY CALLS
		ninehdr(@npage,xdate)

		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		if lSome2Print                           /// return if still false
		 	 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		endif

		xdate:=Date()-240
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-360
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-480
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-600
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-720
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-840
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-960
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-1080
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-1200
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF


		xdate:=Date()-1320
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF


		xdate:=Date()-1440
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-1560
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-1680
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-1800
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-1920
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-2040
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF

		xdate:=Date()-2160
		_CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.t.,aATTR,aLATTR,aNATTR,@nXcoord)
		ENDIF
		///dcprint eject                                /// eject for next report
	ENDIF

	//// now look for Anniversary Dates
	IF lAnnivers
	 xdate:=Date()-365
	 ninehdr(@npage,xdate,.F.)
	 _CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
		if lSome2Print                           /// return if still false
		 	 nintydaycall(xdate,cEmpID,@npage,.F.,aATTR,aLATTR,aNATTR,@nXcoord)
		endif

	 xdate:=Date()-730
	 _CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
	 IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.F.,aATTR,aLATTR,aNATTR,@nXcoord)
	 ENDIF

	 xdate:=Date()-1095
	 _CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
	 IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.F.,aATTR,aLATTR,aNATTR,@nXcoord)
	 ENDIF

	 xdate:=Date()-1460
	 _CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
	 IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.F.,aATTR,aLATTR,aNATTR,@nXcoord)
	 ENDIF

	 xdate:=Date()-1825
	 _CHKFORSOME2PRNT(xdate,@lSome2print,cEmpID)
	 IF lSome2Print
		 nintydaycall(xdate,cEmpID,@npage,.F.,aATTR,aLATTR,aNATTR,@nXcoord)
	 ENDIF
	 ///DCPRINT EJECT
  ENDIF
  /// SKIP 2 LINES
  SKIPALINE()
  IF DCPAGEEJECT()
	 ninehdr(@npage,xdate,.t.)
  ENDIF
  SKIPALINE()
  IF DCPAGEEJECT()
	 ninehdr(@npage,xdate,.t.)
  ENDIF
  IF lInservice
	///dcprint eject
	/// now for current service customers
	xdate:=date()
	ninehdr(@npage,xdate,.t.)
	inservicecall(xdate,cEmpID,@nPage,aATTR,aLATTR,aNATTR,@nXcoord)
	///dcprint eject
  ENDIF
  IF lExpires
	/// now look for expiring leases
	premarklooksm(xdate,cEmpID,@nPAGE,aATTR,aLATTR,aNATTR,@nXcoord)
  ENDIF

  IF lBirthday
	_custbirthday(xdate,cEmpID,@nPAGE,aATTR,aLATTR,aNATTR,@nXcoord)
  ENDIF

  IF nDefmode > 2
	dc_impl(odialog)
  ENDIF

RETURN NIL

static function _chkforsome2prnt(xDate,lSome2print,cEmpID)
lSome2Print:=.f.
select SERVCUST
DBGOTOP()
seek cEmpID
if !found()
 return nil
endif
	/// first check to see if there are any 90 day calls to print BEFORE PRINTING HEADER
do while SERVCUST->sm=cEmpID
 if SERVCUST->deldate=xdate
	if empty(SERVCUST->massmail)
		lSome2Print:=.t.                    /// make true if some are found
	endif
 endif
 skip alias SERVCUST
enddo
RETURN lSome2Print

/// print leases and retails ready to expire in 4 months
static function premarklooksm(xdate,cEmpID,nPAGE,aATTR,aLATTR,aNATTR,nXcoord)
	local lPrevue , dExdate

	premarkhdr(@nPAGE)

	nXCoord:=DC_PRINTERROW()+1
	@nXCoord,0,nXCoord+1.3,131 DCPRINT BOX LINEATTR ALATTR LINEWIDTH 1 FILL GRA_OUTLINEFILL AREAATTR AATTR
	@ dc_printerrow()+1,45 DCPRINT SAY 'Leases/Retails Soon to Expire '

	select SERVCUST
	ordsetfocus('servsm')
	SERVCUST->(dbgotop())
	seek cEmpID
	do while SERVCUST->sm=cEmpID
		if !empty(SERVCUST->deldate)
			dExdate:=SERVCUST->deldate+(30*SERVCUST->term)
			if dExdate - date() > 0
				if dExdate - date() < 120 .AND. dExdate - DATE() > 115
					prnttherec(@nPAGE)
				endif
			endif
		endif
		skip alias SERVCUST
	enddo
return nil



STATIC function nintydaycall(xdate,cEmpID,nPage,lNinety,aATTR,aLATTR,aNATTR,nXCoord)
	select SERVCUST
	DBGOTOP()
	seek cEmpID
	if lninety
		nXCoord:=DC_PRINTERROW()+1
		@nXCoord,0,nXCoord+2.3,131 DCPRINT BOX LINEATTR ALATTR LINEWIDTH 1 FILL GRA_OUTLINEFILL AREAATTR AATTR
		@ DC_PRINTERROW()+2,44 DCPRINT SAY 'One Hundred Twenty Day Calls For  '+ dtoc(xdate)
	else
		nXCoord:=DC_PRINTERROW()+1
		@nXCoord,0,nXCoord+2.3,131 DCPRINT BOX LINEATTR ALATTR LINEWIDTH 1 FILL GRA_OUTLINEFILL AREAATTR AATTR
		@ DC_PRINTERROW()+2,44 DCPRINT SAY 'Anniverary Cards for Customers Delivered. '+ dtoc(xdate)
	endif
	do while SERVCUST->sm=cEmpID
		if SERVCUST->deldate=xdate
			if empty(SERVCUST->massmail)           /// DO NOT  mail if anything in this field
				@ DC_PRINTERROW()+1,0 DCPRINT SAY ' '
				if dcpageeject()
					ninehdr(@npage,xdate)
				endif
				Prnt90day(TOP_MAR(),BOT_MAR(),PAGE_LEN(),@nPAGE,xdate)
			endif
		endif                                     /// deldate=xdate
		skip alias SERVCUST
	enddo
return nil




STATIC function inservicecall(xdate,cEmpID,nPage,aATTR,aLATTR,aNATTR,nXcoord)
	local lSome2Print:=.f.
	select SERVCUST
	DBGOTOP()
	seek cEmpID
	if !found()
		return nil
	endif
	/// first check to see if there are any service calls to print
	do while SERVCUST->sm=cEmpID
		if SERVCUST->lastserv=xdate.or.SERVCUST->lastserv=xdate-1
			if empty(SERVCUST->massmail)
				lSome2Print:=.t.                    /// make true if some are found
			endif
		endif
		skip alias SERVCUST
	enddo
	if lSome2Print=.f.                           /// return if still false
		return nil
	endif

	/// if some are found
	select SERVCUST
	DBGOTOP()
	seek cEmpID
	SKIPALINE()
	nXCoord:=DC_PRINTERROW()+1
	@nXCoord,0,nXCoord+1.3,131 DCPRINT BOX LINEATTR ALATTR LINEWIDTH 1 FILL GRA_OUTLINEFILL AREAATTR AATTR
	@ DC_PRINTERROW()+1,40 DCPRINT SAY 'Your Customers Visiting Service Calls For  '+ dtoc(xdate)
	do while SERVCUST->sm=cEmpID
		if SERVCUST->lastserv=xdate.or.SERVCUST->lastserv=xdate-1
			if empty(SERVCUST->massmail)           /// DO NOT  mail if anything in this field
				@ DC_PRINTERROW()+1,0 DCPRINT SAY ' '
				if dcpageeject()
					ninehdr(@npage,xdate,.t.)
				endif
				Prnt90day(TOP_MAR(),BOT_MAR(),PAGE_LEN(),@nPAGE,xdate,.t.)
			endif
		endif                                     /// deldate=xdate
		skip alias SERVCUST
	enddo
	dcprint eject
	/// now look for upcoming appointments TO DECIDE whether to print or not
	lSome2print:=.f.
	SELECT SERVRESV
	SEEK cEmpID
	DO WHILE SERVRESV->SALESMAN=cEmpID
  		 IF SERVRESV->RESDATE=xdate .OR. SERVRESV->RESDATE=xdate+1
  			  lSome2print:=.t.
  		     EXIT
		 ENDIF
  		 SKIP ALIAS SERVRESV
   ENDDO
	// IF THERE ARE ANY - PRINT THEM
	IF lSome2print
		 nXCoord:=DC_PRINTERROW()+1
	    @nXCoord,0,nXCoord+1.3,131 DCPRINT BOX LINEATTR ALATTR LINEWIDTH 1 FILL GRA_OUTLINEFILL AREAATTR AATTR
	    @ DC_PRINTERROW()+1,40 DCPRINT SAY 'Your Customers With UpComing Service Appointments For  '+ dtoc(xdate)+'  -  '+dtoc(xdate+1)

		_PLNAPPTHDR(@nPAGE,xDate)

		SELECT SERVRESV
		SERVRESV->(DBGOTOP())
	   seek cEmpID
	   DO WHILE SERVRESV->SALESMAN=cEmpID
		  IF SERVRESV->RESDATE=xdate .OR. SERVRESV->RESDATE=xdate+1
					_PrntDayAppts(TOP_MAR(),BOT_MAR(),PAGE_LEN(),@nPAGE,xdate)
					IF DCPAGEEJECT()
						_PLNAPPTHDR(@nPAGE,xDate)
					ENDIF
		   ENDIF
		   SKIP ALIAS SERVRESV
		ENDDO
	dcprint eject
	ENDIF

return nil



static function _PrntDayAppts(nPAGE,xdate)
	@ DC_PRINTERROW()+1,0 DCPRINT SAY SERVRESV->RESDATE
	if servresv->waityn='Y'
		@ DC_PRINTERROW(),12 DCPRINT SAY SERVRESV->TEAM font '8.Courier New Bold'
		@ DC_PRINTERROW(),15 DCPRINT SAY SERVRESV->LASTNAME font '8.Courier New Bold'
		@ DC_PRINTERROW(),36 DCPRINT SAY SERVRESV->FIRSTNAME font '8.Courier New Bold'
	else
		@ DC_PRINTERROW(),12 DCPRINT SAY SERVRESV->TEAM
		@ DC_PRINTERROW(),15 DCPRINT SAY SERVRESV->LASTNAME
		@ DC_PRINTERROW(),36 DCPRINT SAY SERVRESV->FIRSTNAME
	endif
	@ DC_PRINTERROW(),60 DCPRINT SAY SERVRESV->YEAR
	@ DC_PRINTERROW(),65 DCPRINT SAY SERVRESV->CARDESC
	@ DC_PRINTERROW(),85 DCPRINT SAY SERVRESV->VIN
	@ DC_PRINTERROW(),107 DCPRINT SAY SERVRESV->APPTTIME font '8.Courier New Bold'
	@ DC_PRINTERROW(),117 DCPRINT SAY SUBSTR(SERVRESV->COMMENT,1,10)
	@ DC_PRINTERROW(),129 DCPRINT SAY SERVRESV->waityn font '8.Courier New Bold'
	IF DCPAGEEJECT()
		_PLNAPPTHDR(@nPAGE,xDate)
	ENDIF
	@ DC_PRINTERROW()+1,0 DCPRINT SAY SERVRESV->HPHONE
	@ DC_PRINTERROW(),20 DCPRINT SAY SERVRESV->WPHONE
	@ DC_PRINTERROW(),40 DCPRINT SAY SERVRESV->EMAIL
	IF DCPAGEEJECT()
		_PLNAPPTHDR(@nPAGE,xDate)
	ENDIF
	@ DC_PRINTERROW()+1,0 DCPRINT SAY SERVRESV->CONCERN
	@ DC_PRINTERROW(),70 DCPRINT SAY SERVRESV->CONCERN2
	IF DCPAGEEJECT()
		_PLNAPPTHDR(@nPAGE,xDate)
	ENDIF
	skipaline()
	IF DCPAGEEJECT()
		_PLNAPPTHDR(@nPAGE,xDate)
	ENDIF

	RETURN NIL

static function _PLNAPPTHDR(nPAGE,xDate)
	nPAGE++
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Up Coming Appointments.  Page '+alltrim(str(npage))  font '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'ApptDate' font '8.Courier New Bold'
	@ DC_PRINTERROW(),10  DCPRINT SAY 'Team'     font '8.Courier New Bold'
	@ DC_PRINTERROW(),15  DCPRINT SAY 'Lastname' font '8.Courier New Bold'
	@ DC_PRINTERROW(),36  DCPRINT SAY 'Firstname' font '8.Courier New Bold'
	@ DC_PRINTERROW(),60  DCPRINT SAY 'Year'   font '8.Courier New Bold'
	@ DC_PRINTERROW(),65  DCPRINT SAY 'Vehicle' font '8.Courier New Bold'
	@ DC_PRINTERROW(),85  DCPRINT SAY 'Vin'  font '8.Courier New Bold'
	@ DC_PRINTERROW(),105  DCPRINT SAY 'Time' font '8.Courier New Bold'
	@ DC_PRINTERROW(),115  DCPRINT SAY 'Mileage' font '8.Courier New Bold'
	@ DC_PRINTERROW(),125  DCPRINT SAY 'Wait'   font '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'HomePhn' font '8.Courier New Bold'
	@ DC_PRINTERROW(),20 DCPRINT SAY 'WorkPhn'  font '8.Courier New Bold'
	@ DC_PRINTERROW(),40 DCPRINT SAY 'eMail Address' font '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Concerns'  font '8.Courier New Bold'
	RETURN NIL

static function _custbirthday(xDate, cEmpID,nPage,aATTR,aLATTR,aNATTR,nXcoord)
local lSome2Print:=.f.
LOCAL cxDate:=space(4),cxDate2:=SPACE(4),cxDate3:=space(4),cxDate4:=SPACE(4)
	select SERVCUST
	DBGOTOP()
	seek cEmpID
	if !found()
		return nil
	endif
	/// first check to see if there are any birthdays to print

	cxDate:=substr(dtos(xdate),5,4)
	cXdATE2:=SUBSTR(DTOS(xDate+1),5,4)
	cXdATE3:=SUBSTR(DTOS(xDate+2),5,4)
	cXdATE4:=SUBSTR(DTOS(xDate+3),5,4)

	do while SERVCUST->sm=cEmpID
		if substr(DTOS(SERVCUST->buyerdob),5,4)=cXdate .or. substr(DTOS(SERVCUST->buyerdob),5,4)=cxDate2 .OR. ;
				substr(DTOS(SERVCUST->buyerdob),5,4)=cxDate3 .OR. substr(DTOS(SERVCUST->buyerdob),5,4)=cxDate4
			if empty(SERVCUST->massmail)
				IF DATE()-SERVCUST->LASTSERV < 730
				 lSome2Print:=.t.                    /// make true if some are found
				ENDIF
			endif
		endif
		skip alias SERVCUST
	enddo
	if lSome2Print=.f.                           /// return if still false
		return nil
	endif

	/// if some are found
	select SERVCUST
	DBGOTOP()
	seek cEmpID
	SKIPALINE()
	nXCoord:=DC_PRINTERROW()+1
	@nXCoord,0,nXCoord+1.3,131 DCPRINT BOX LINEATTR ALATTR LINEWIDTH 1 FILL GRA_OUTLINEFILL AREAATTR AATTR
	@ DC_PRINTERROW()+1,20 DCPRINT SAY 'Your Customers Birthdays For   '+ dtoc(xdate) + ' and ' + dtoc(xdate+1)+ ' and ' + dtoc(xdate+2);
		+ ' and ' + dtoc(xdate+3)

	birthdatehdr(@npage,xdate)

	do while SERVCUST->sm=cEmpID

		if substr(DTOS(SERVCUST->buyerdob),5,4)=cXdate.or.substr(DTOS(SERVCUST->buyerdob),5,4)=cxDate2 .OR. ;
				substr(DTOS(SERVCUST->buyerdob),5,4)=cxDate3 .OR. substr(DTOS(SERVCUST->buyerdob),5,4)=cxDate4
			if empty(SERVCUST->massmail)
				IF DATE()-SERVCUST->LASTSERV < 730
		   		PrntBirthday(TOP_MAR(),BOT_MAR(),PAGE_LEN(),@nPAGE,xdate)
				ENDIF
			endif
		endif                                     /// deldate=xdate
		skip alias SERVCUST
	enddo
	dcprint eject
RETURN NIL








//// function to print 90 day call lines
STATIC Function Prnt90day(nPAGE,xdate,inserv)
	default inserv:=.f.
	@ DC_PRINTERROW()+1,0 DCPRINT SAY SERVCUST->SM
	@ DC_PRINTERROW(),5 DCPRINT SAY SERVCUST->FIRSTNAME
	@ DC_PRINTERROW(),31 DCPRINT SAY SERVCUST->LASTNAME
	@ DC_PRINTERROW(),52 DCPRINT SAY SERVCUST->phone  FONT '8.Courier New Bold'
	@ DC_PRINTERROW(),70 DCPRINT SAY SERVCUST->make
	@ DC_PRINTERROW(),90 DCPRINT SAY SERVCUST->cardesc
	if dcpageeject()
		ninehdr(@npage,xdate,inserv)
	endif
	@ DC_PRINTERROW()+1,0 DCPRINT SAY ' '
	@ DC_PRINTERROW(),5 DCPRINT SAY SERVCUST->STREET
	@ DC_PRINTERROW(),31 DCPRINT SAY SERVCUST->lastserv
	@ DC_PRINTERROW(),42 DCPRINT SAY ' / '+alltrim(str(SERVCUST->lastmiles))
	@ DC_PRINTERROW(),60 DCPRINT SAY SERVCUST->deldate
	if dcpageeject()
		ninehdr(@npage,xdate,inserv)
	endif

	@ DC_PRINTERROW()+1,0 DCPRINT SAY '  '
	@ DC_PRINTERROW(),5 DCPRINT SAY SERVCUST->CITYSTATE
	@ DC_PRINTERROW(),31 DCPRINT SAY SERVCUST->ZIP
	@ DC_PRINTERROW(),52 DCPRINT SAY SERVCUST->EMAIL
	if dcpageeject()
		ninehdr(@npage,xdate,inserv)
	endif

return nil

//// function to print Birthday call lines
STATIC Function PrntBirthday(nPAGE,xdate)
	@ DC_PRINTERROW()+1,0 DCPRINT SAY SERVCUST->SM
	@ DC_PRINTERROW(),5 DCPRINT SAY SERVCUST->FIRSTNAME
	@ DC_PRINTERROW(),31 DCPRINT SAY SERVCUST->LASTNAME
	@ DC_PRINTERROW(),52 DCPRINT SAY SERVCUST->phone  FONT '8.Courier New Bold'
	@ DC_PRINTERROW(),70 DCPRINT SAY SERVCUST->make
	@ DC_PRINTERROW(),90 DCPRINT SAY SERVCUST->cardesc
	if dcpageeject()
		birthdatehdr(@npage,xdate)
	endif
	@ DC_PRINTERROW()+1,0 DCPRINT SAY ' '
	@ DC_PRINTERROW(),5 DCPRINT SAY SERVCUST->STREET
	@ DC_PRINTERROW(),31 DCPRINT SAY SERVCUST->lastserv
	@ DC_PRINTERROW(),42 DCPRINT SAY ' / '+alltrim(str(SERVCUST->lastmiles))
	@ DC_PRINTERROW(),60 DCPRINT SAY SERVCUST->buyerdob FONT '10.Tahoma Bold'

	if dcpageeject()
		birthdatehdr(@npage,xdate)
	endif

	@ DC_PRINTERROW()+1,0 DCPRINT SAY '  '
	@ DC_PRINTERROW(),5 DCPRINT SAY SERVCUST->CITYSTATE
	@ DC_PRINTERROW(),31 DCPRINT SAY SERVCUST->ZIP
	@ DC_PRINTERROW(),52 DCPRINT SAY SERVCUST->EMAIL
	if dcpageeject()
		birthdatehdr(@npage,xdate)
	endif

return nil



static function ninehdr(npage,xdate,inserv)
	DEFAULT inserv:=.f.
	nPage:=npage+1
	if inserv=.t.
		@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Customers in Service  '+ 'Date in '+dtoc(xdate)+' to '+dtoc(xdate-1)+ '    Page '+ alltrim(str(npage))     FONT '8.Courier New Bold'
	else
		@ DC_PRINTERROW()+1,0 DCPRINT SAY '120 day calls  '+ 'Delivered '+dtoc(xdate)+ '    Page '+ alltrim(str(npage))     FONT '8.Courier New Bold'
	endif
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'SM   FirstName                       LastName                   Phone    Make   Desc'  FONT '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY '     Street                      Last Service/Mileage           Delivery Date '        FONT '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY '     City/State                       Zip                       Email Address'                               FONT '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY ' '
return nil

static function Birthdatehdr(npage,xdate)
	nPage:=npage+1
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Customers Birthdates  '+ dtoc(xdate)+' to '+dtoc(xdate-1)+ '    Page '+ alltrim(str(npage))     FONT '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'SM   FirstName                       LastName                   Phone    Make   Desc'  FONT '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY '     Street                      Last Service/Mileage           BIRTHDATE '        FONT '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY '     City/State                       Zip                       Email Address'                               FONT '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY ' '
return nil







STATIC Function PrntPlan(nPAGE,lDel)
	LOCAL nRec:=0
	/// return if upstat is F and it is not within 10 days
	if NEWUPS->upstatus='F'
		// if not within 3 days return
		if abs(NEWUPS->tickle-date()) > 1
			return nil
		endif
		/// if within 3 days make it active
		if abs(NEWUPS->tickle-Date()) < 1
			if NEWUPS->(dc_reclock())
				nRec:=NEWUPS->(RECNO())
				replace NEWUPS->upstatus with 'A'
				NEWUPS->(dbcommit())
				NEWUPS->(dbrunlock(nRec))
			endif
		endif
	endif
	@ DC_PRINTERROW()+1,0 DCPRINT SAY NEWUPS->SALESMAN font '8.Courier New Bold'
	@ DC_PRINTERROW(),5 DCPRINT SAY SUBSTR(NEWUPS->FIRSTNAME,1,20)
	@ DC_PRINTERROW(),31 DCPRINT SAY SUBSTR(NEWUPS->LASTNAME,1,20)
	@ DC_PRINTERROW(),52 DCPRINT SAY NEWUPS->AREA   font '8.Courier New Bold'
	@ DC_PRINTERROW(),63 DCPRINT SAY NEWUPS->UPID   font '8.Courier New Bold'
	@ DC_PRINTERROW(),72 DCPRINT SAY NEWUPS->COMMENTS
	@ DC_PRINTERROW()+1,0 DCPRINT SAY '  '
	if dcpageeject()
		planhdr(@npage)
	endif
	@ DC_PRINTERROW(),5 DCPRINT SAY SUBSTR(NEWUPS->STREET,1,20)
	@ DC_PRINTERROW(),31 DCPRINT SAY NEWUPS->UPTYPE  font '8.Courier New Bold'
	@ DC_PRINTERROW(),33 DCPRINT SAY NEWUPS->INTEREST
	@ DC_PRINTERROW(),48 DCPRINT SAY NEWUPS->DATEIN
	@ DC_PRINTERROW(),60 DCPRINT SAY NEWUPS->TICKLE  font '8.Courier New Bold'
	@ DC_PRINTERROW(),72 DCPRINT SAY NEWUPS->COMMENT2
	@ DC_PRINTERROW()+1,0 DCPRINT SAY '  '
	if dcpageeject()
		planhdr(@npage)
	endif
	@ DC_PRINTERROW(),5 DCPRINT SAY NEWUPS->CITYSTATE
	@ DC_PRINTERROW(),31 DCPRINT SAY NEWUPS->ZIPCODE
	@ DC_PRINTERROW(),40 DCPRINT SAY NEWUPS->PREFMETHOD
	@ DC_PRINTERROW(),52 DCPRINT SAY NEWUPS->EMAIL   font '8.Courier New Bold'
	@ dc_printerrow(),100 DCPRINT SAY NEWUPS->Upstatus font '8.Courier New Bold'
	@ dc_printerrow(),103 DCPRINT SAY NEWUPS->Tomgr font '8.Courier New Bold'
	if ldel=.t.                                  /// print deliveries
		@ dc_printerrow(),110 DCPRINT SAY NEWUPS->Billstat font '8.Courier New Bold'
	   @ DC_PRINTERROW(),115 DCPRINT SAY NEWUPS->SENDLTR
	endif
return nil


static function planhdr(npage)
	nPage:=npage+1
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Daily Planner Page  '+alltrim(str(npage))      FONT '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'SM    FirstName              LastName           Area           Phone   ID    Comments' FONT '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY '      Street                 Uptype  Interest  Date In   Follow Date ' FONT '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY '      City/State                   Zip      PrefMeth   Email Address       Status  TOMgr  DelStat  DelFollow                                 Upstat  ToMgr   BillStat'   FONT '8.Courier New Bold'
	@ DC_PRINTERROW()+1,0 DCPRINT SAY ' '        /// SKIP LINE
return nil

static function allplans(dCutoff,lAnnivers,lOneTwenty,lDeliveries,lExpires,lInservice,lBirthday,nDefmode)
	local lFoundups:=.t.
	select sm
	goto top
	do while !sm->(eof())

		if sm->ACTIVE='Y'
			oneplan(sm->salesman,@lFoundups,dCutoff,lAnnivers,lOneTwenty,lDeliveries,lExpires,lInservice,lBirthday,nDefmode)

			if lFoundups
				dcprint eject
			endif
		endif

		skip alias sm
	enddo
return nil

static Function _Optionsoft(nRec)
	local cBody:=''
	local cOptPath:='c:\OPTIONSOFT\DEALS\PENDING\' ,cLastname
	select newups
	NEWUPS->(dbgoto(nRec))
	_makehdr(@cBody)     /// make schema section
	/// now make data section
	cBody:=cBody+'<dealtype>'+CRLF
	IF NEWUPS->NEWUSED='N'
    cBody:=cBody+'<type>NewCar</type>'+CRLF
	ELSE
	 cBody:=cBody+'<type>UsedCar</type>'+CRLF
	ENDIF
	 cBody:=cBody+'<status>Pending</status>'+CRLF
    cBody:=cBody+'<dealnum />'+CRLF
    cBody:=cBody+'</dealtype>'+CRLF
    cBody:=cBody+'<users>'+CRLF
    cBody:=cBody+'<customer>'+alltrim(NEWUPS->LASTNAME)+'</customer>'+CRLF
    cBody:=cBody+'<manager>0</manager>'+CRLF
    cBody:=cBody+'<salesman>-1</salesman>'+CRLF
    cBody:=cBody+'</users>'+CRLF
    cBody:=cBody+'<vehicle>'+CRLF
    cBody:=cBody+'<stocknum>'+NEWUPS->STKQUOTED+'</stocknum>'+CRLF
    cBody:=cBody+'<sellprice>'+alltrim(str(NEWUPS->QUOTED))+'</sellprice>'+CRLF
    cBody:=cBody+'<trade>'+alltrim(str(NEWUPS->TRADEQUOTE)) +'</trade>'+CRLF
    cBody:=cBody+'<payoff>'+alltrim(str(NEWUPS->LIEN)) +'</payoff>'+CRLF
    cBody:=cBody+'<rebate>'+alltrim(str(NEWUPS->REBATE)) +'</rebate>'+CRLF
    cBody:=cBody+'<cashdown>'+alltrim(str(NEWUPS->DOWNPAY)) +'</cashdown>'+CRLF
    cBody:=cBody+'</vehicle>'+CRLF
    cBody:=cBody+'<rates>'+CRLF
    cBody:=cBody+'<monthone>36</monthone>'+CRLF
    cBody:=cBody+'<aprone>7.99</aprone>'+CRLF
    cBody:=cBody+'<monthtwo>48</monthtwo>'+CRLF
    cBody:=cBody+'<aprtwo>7.99</aprtwo>'+CRLF
    cBody:=cBody+'<monththree>60</monththree>'+CRLF
    cBody:=cBody+'<aprthree>7.99</aprthree>'+CRLF
    cBody:=cBody+'<monthfour>72</monthfour>'+CRLF
    cBody:=cBody+'<aprfour>8.99</aprfour>'+CRLF
    cBody:=cBody+'</rates>'+CRLF
    cBody:=cBody+'<vsc>'+CRLF
    cBody:=cBody+'<months>99</months>'+CRLF
    cBody:=cBody+'<miles>100000</miles>'+CRLF
    cBody:=cBody+'<deduct>$250.00</deduct>'+CRLF
    cBody:=cBody+'</vsc>'+CRLF
    cBody:=cBody+'<fees>'+CRLF
    cBody:=cBody+'<doc>'+alltrim(str(NEWUPS->PROCFEE))+'</doc>'+CRLF
    cBody:=cBody+'<dmv>'+alltrim(str(NEWUPS->REGFEE))+'</dmv>'+CRLF
    cBody:=cBody+'<othertax>0</othertax>'+CRLF
    cBody:=cBody+'<salestax>'+alltrim(str(NEWUPS->TAXPER))+'</salestax>'+CRLF
    cBody:=cBody+'<vsi>22.5</vsi>'+CRLF
    cBody:=cBody+'<countytax>0</countytax>'+CRLF
    cBody:=cBody+'</fees>'+CRLF
    cBody:=cBody+'<payment>'+CRLF
    cBody:=cBody+'<term>72</term>'+CRLF
    cBody:=cBody+'<daystopay>45</daystopay>'+CRLF
  cBody:=cBody+'</payment>'+CRLF
cBody:=cBody+'</deal>'+CRLF

 cLastName:=alltrim(NEWUPS->LASTNAME)
/// write xml file
DCPRINT ON TEXTONLY OUTFILE 'c:\OPTIONSOFT\DEALS\PENDING\'+cLastName+'_Pending.xml' OVERWRITE
          DCPRINT ?? cBody
          DCPRINT OFF

			 DC_WINALERT('FILE CREATED '+cLastName)

return nil






static Function _Makehdr(cBody)
cBody:='<?xml version="1.0" standalone="yes"?>'+CRLF
    cBody:=cBody+'<deal>'+CRLF
    cBody:=cBody+'<xs:schema id="deal" xmlns="" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:msdata="urn:schemas-microsoft-com:xml-msdata">'+CRLF
     cBody:=cBody+'<xs:element name="deal" msdata:IsDataSet="true" msdata:UseCurrentLocale="true">'+CRLF
      cBody:=cBody+'<xs:complexType>'+CRLF
        cBody:=cBody+'<xs:choice minOccurs="0" maxOccurs="unbounded">'+CRLF
          cBody:=cBody+'<xs:element name="Deal">'+CRLF
            cBody:=cBody+'<xs:complexType>'+CRLF
              cBody:=cBody+'<xs:sequence>'+CRLF
                cBody:=cBody+'<xs:element name="dealtype" minOccurs="0" maxOccurs="unbounded">'+CRLF
                  cBody:=cBody+'<xs:complexType>'+CRLF
                    cBody:=cBody+'<xs:sequence>'+CRLF
                      cBody:=cBody+'<xs:element name="type" type="xs:string" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="status" type="xs:string" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="dealnum" type="xs:string" />'+CRLF
                    cBody:=cBody+'</xs:sequence>'+CRLF
                  cBody:=cBody+'</xs:complexType>'+CRLF
                cBody:=cBody+'</xs:element>'+CRLF
                cBody:=cBody+'<xs:element name="users" minOccurs="0" maxOccurs="unbounded">'+CRLF
                  cBody:=cBody+'<xs:complexType>'+CRLF
                    cBody:=cBody+'<xs:sequence>'+CRLF
                      cBody:=cBody+'<xs:element name="customer" type="xs:string" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="manager" type="xs:int" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="salesman" type="xs:int" minOccurs="0" />'+CRLF
                    cBody:=cBody+'</xs:sequence>'+CRLF
                  cBody:=cBody+'</xs:complexType>'+CRLF
                cBody:=cBody+'</xs:element>'+CRLF
                cBody:=cBody+'<xs:element name="vehicle" minOccurs="0" maxOccurs="unbounded">'+CRLF
                  cBody:=cBody+'<xs:complexType>'+CRLF
                    cBody:=cBody+'<xs:sequence>'+CRLF
                      cBody:=cBody+'<xs:element name="stocknum" type="xs:string" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="sellprice" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="trade" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="payoff" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="rebate" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="cashdown" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="capcost" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="msrp" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="adminfee" type="xs:double" minOccurs="0" />'+CRLF
                    cBody:=cBody+'</xs:sequence>'+CRLF
                  cBody:=cBody+'</xs:complexType>'+CRLF
                cBody:=cBody+'</xs:element>'+CRLF
                cBody:=cBody+'<xs:element name="rates" minOccurs="0" maxOccurs="unbounded">'+CRLF
                  cBody:=cBody+'<xs:complexType>'+CRLF
                    cBody:=cBody+'<xs:sequence>'+CRLF
                      cBody:=cBody+'<xs:element name="resone" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="moneyone" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="monthone" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="aprone" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="restwo" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="moneytwo" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="monthtwo" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="aprtwo" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="resthree" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="moneythree" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="monththree" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="aprthree" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="resfour" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="moneyfour" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="monthfour" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="aprfour" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="useamt" type="xs:boolean" minOccurs="0" />'+CRLF
                    cBody:=cBody+'</xs:sequence>'+CRLF
                  cBody:=cBody+'</xs:complexType>'+CRLF
                cBody:=cBody+'</xs:element>'+CRLF
                cBody:=cBody+'<xs:element name="vsc" minOccurs="0" maxOccurs="unbounded">'+CRLF
                  cBody:=cBody+'<xs:complexType>'+CRLF
                    cBody:=cBody+'<xs:sequence>'+CRLF
                      cBody:=cBody+'<xs:element name="months" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="miles" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="deduct" type="xs:string" minOccurs="0" />'+CRLF
                    cBody:=cBody+'</xs:sequence>'+CRLF
                  cBody:=cBody+'</xs:complexType>'+CRLF
                cBody:=cBody+'</xs:element>'+CRLF
                cBody:=cBody+'<xs:element name="fees" minOccurs="0" maxOccurs="unbounded">'+CRLF
                  cBody:=cBody+'<xs:complexType>'+CRLF
                    cBody:=cBody+'<xs:sequence>'+CRLF
                      cBody:=cBody+'<xs:element name="doc" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="dmv" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="othertax" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="salestax" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="vsi" type="xs:double" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="countytax" type="xs:double" minOccurs="0" />'+CRLF
                    cBody:=cBody+'</xs:sequence>'+CRLF
                  cBody:=cBody+'</xs:complexType>'+CRLF
                cBody:=cBody+'</xs:element>'+CRLF
                cBody:=cBody+'<xs:element name="payment" minOccurs="0" maxOccurs="unbounded">'+CRLF
                  cBody:=cBody+'<xs:complexType>'+CRLF
                    cBody:=cBody+'<xs:sequence>'+CRLF
                      cBody:=cBody+'<xs:element name="term" type="xs:int" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="daystopay" type="xs:int" minOccurs="0" />'+CRLF
                    cBody:=cBody+'</xs:sequence>'+CRLF
                  cBody:=cBody+'</xs:complexType>'+CRLF
                cBody:=cBody+'</xs:element>'+CRLF
                cBody:=cBody+'<xs:element name="products" minOccurs="0" maxOccurs="unbounded">'+CRLF
                  cBody:=cBody+'<xs:complexType>'+CRLF
                    cBody:=cBody+'<xs:sequence>'+CRLF
                      cBody:=cBody+'<xs:element name="id" type="xs:int" minOccurs="0" />'+CRLF
                      cBody:=cBody+'<xs:element name="amount" type="xs:double" minOccurs="0" />'+CRLF
                    cBody:=cBody+'</xs:sequence>'+CRLF
                  cBody:=cBody+'</xs:complexType>'+CRLF
                cBody:=cBody+'</xs:element>'+CRLF
              cBody:=cBody+'</xs:sequence>'+CRLF
            cBody:=cBody+'</xs:complexType>'+CRLF
          cBody:=cBody+'</xs:element>'+CRLF
        cBody:=cBody+'</xs:choice>'+CRLF
      cBody:=cBody+'</xs:complexType>'+CRLF
    cBody:=cBody+'</xs:element>'+CRLF
  cBody:=cBody+'</xs:schema>'+CRLF
  RETURN cBody


static Function GetOptionsoft(cFile,aNums)
	local getlist:={}, i,cId,cAmt
	local oOptionMain,oDeal,oProducts,oID,oAmt,nSoft1,nSoft2,nSoft3,nSoft4,nSoft5,nSoft6,oVehicle,oCashdown,nCashdown:=0,cCashdown
	Local nVsc:=0


	nSoft1:=nSoft2:=nSoft3:=nSoft4:=nSoft5:=nSoft6:=0
	oOptionmain:=dc_xml2objectTree(cFile)
	IF valtype(oOptionmain) <> 'O'
		dc_winalert('No File Selected')
		return nil
	ENDIF
	oDeal:=oOptionmain:findnode('deal')
	oProducts:=oDeal:findnode('products',.t.)
	FOR i := 1 TO len(oProducts)
		oID:=oProducts[i]:findnode('id')
		cId:=oID:content
		oAMT:=oProducts[i]:findnode('amount')
		cAmt:=oAmt:content
		IF cID='1'
			nVsc:=val(cAmt)
		endif
		IF cID='2'
			nSoft1:=val(cAmt)
		endif
		IF cID='3'
			nSoft2:=val(cAmt)
		endif
		IF cID='4'
			nSoft3:=val(cAmt)
		endif
		IF cID='5'
			nSoft4:=val(cAmt)
		endif
		IF cID='6'
			nSoft5:=val(cAmt)
		endif
		IF cID='7'
			nSoft6:=val(cAmt)
		endif
	next
	aNums[QUOTEN_NAFTSALEAM1]:=nSoft1
   aNums[QUOTEN_NAFTSALEAM2]:=nSoft2
   aNums[QUOTEN_NAFTSALEAM3]:=nSoft3
   aNums[QUOTEN_NAFTSALEAM4]:=nSoft4
   aNums[QUOTEN_NAFTSALEAM5]:=nSoft5
   aNums[QUOTEN_NAFTSALEAM6]:=nSoft6
	aNums[QUOTEN_NSERVCONAM]:=nVsc

	/*
	oVehicle:=oDeal:FINDNODE('vehicle')
	oCashdown:=oVehicle:Findnode('cashdown')
	cCashdown:=oCashdown:content
	nCashdown:=val(cCashdown)
	aNums[QUOTEN_NDOWNPAY]:=nCashdown  */

	return nil


FUNCTION UPQUOTE(nUPRec)
	local getlist:={},oRecord ,xstatus, cProgram ,nCopies:=1 ,ACon ,lWarn:=.F.,oRebate,oFinance,oTrade,;
		oTabpage0,oTabpage1,oTabpage2,oTabpage3,oTabpage4,oTabpage5,oTabpageCalcs,oLease,oAftsale,oDealinfo
	local getoptions , oToolbar,oToolbar2,lOfferOC:=.f. , dDelivdate , cDeltime, nRec:=0, cBody,nProfit, cFilename:='' ,oFees, nTotnonveh:=0 ,lChng:=.f.
	local clLoadPayment:='N',nLoad:=0,pf:=0 ,oLRebate,oToolbar3,oToolbar4,oToolbar5, lRecalc:=.f.
	LOCAL lApptoLs1:=.f.,lApptoLs2:=.f.,lApptoLs3:=.f.,lApptoLs4:=.f. ,lApptoLs5:=.f. , lApptoLs6:=.f., oRebateY
	LOCAL AoToolbar,nCurrUptype,aPres,oBrowse,aAppts:={},lDelivered:=.f. ,lCaptax:=.t.,lCapfees:=.t.,lProptax:=.f.
	Local oConfig1,oConfig2 ,oUpinfo,oVehinfo,oUpstat,oProgInfo,oProgInfo2,oRetailCalcs,oToolCalc ,oLeasecalcs ,oLeasesum,oLRebatex
	DCGETOPTIONS saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize NOTABSTOP  NOSUPERVISE
	cProgram:='UPQUOTE'

	/// got to record and create record object
	NEWUPS->(DBGOTO(nUPRec))
	oRecord := NEWUPS->(DC_DbRecord():new())
   NEWUPS->(Db_Scatter(@oRecord))

	///plug a status if it is empty
	IF EMPTY(oRecord:status)
		oRecord:status='A'
	ENDIF
 	oRecord:addtopymnt:=0

	IF oRecord:lpayment2 > 0
		lRecalc:=.t.
	ENDIF

	oConfig1 := DC_XbpPushButtonXPConfig():new()
	oConfig1:bitmapOffset := 5
	oConfig1:fgColorMouse := COLOR_BLACK
	oConfig1:bgColorMouse := COLOR_SILVER
	oConfig1:fgColor := COLOR_BLACK
	oConfig1:bgColor := BD_CORNFLOWERBLUE
	oConfig1:gradientStep := 10
	oConfig1:gradientReverse := .t.
	oConfig1:radius := 10

	oConfig2 := DC_XbpPushButtonXPConfig():new()
	oConfig2:bitmapOffset := 5
	oConfig2:fgColorMouse := COLOR_BLACK
	oConfig2:bgColorMouse := COLOR_YELLOW
	oConfig2:fgColor := COLOR_DARKBLUE
	oConfig2:bgColor := COLOR_WHITE
	oConfig2:gradientStep := 10
	oConfig2:gradientReverse := .t.
	oConfig2:radius := 10


	IF EMPTY(oRecord:UPID)
		DC_WINALERT('You must have an UP ID to maintain. Please put one in on the browser.')
      RETURN NIL
	ENDIF

	IF oRecord:Billstat='D'
	 	lDelivered:=.T.
	ENDIF

	nCurrUpType:=NEWUPS->UPTYPE       /// set up to look for a change


IF.NOT.SECURITY(@cProgram)
	DC_winALERT('SECURITY VIOLATION')
   RETURN NIL
ENDIF


IF !UseDb({'UPTAXTAB','FNIOPTS','SERVCONT'})
	 DC_WINALERT('Cannot Open Files     .')
    RETURN NIL
ENDIF



nProfit:=0
cBody:=''
dDelivdate:=NEWUPS->Delivdate
cDeltime:=NEWUPS->Deltime
nTotNonVeh:=0

/// ADD UP NON VEHICLE COSTS SO FAR
nTotNonVeh:=oRecord:AFTSALEAM1+oRecord:AFTSALEAM2+oRecord:AFTSALEAM3+oRecord:AFTSALEAM4;
	+oRecord:AFTSALEAM5+oRecord:AFTSALEAM6+oRecord:INSPECTION+oRecord:REGFEE+oRecord:PROCFEE;
	+oRecord:WASTEFEE+oRecord:SERVCONAM+oRecord:AFTSALETAX

oRecord:PROCFEE := FNIOPTS->Docfee

IF oRecord:INTRATE < .01
	oRecord:INTRATE := 4.9
ENDIF
IF oRecord:TERMODD=0
	oRecord:TERMODD:=84
ENDIF

IF oRecord:captax='Y'
	lCaptax:=.t.
else
	lCaptax:=.f.
ENDIF

IF oRecord:capfees='Y'
	lCapfees:=.t.
else
	lCapfees:=.f.
ENDIF

IF oRecord:extrachar3='Y'
	lProptax:=.t.
else
	lProptax:=.f.
ENDIF

/// set initial value
oRecord:aprlease:=.f.






@ 0,0 DCTABPAGE oTabpage0 CAPTION 'Prospect Information' ;
		size 150,34  ;

   @ 0,1 DCGROUP oUpStat Caption 'Prospect Status' Size 140,7.2  PARENT oTabpage0

	@ 1,1 DCSAY 'Up Id' GET oRecord:UPID  SAYCOLOR 1,BD_LEMONCREME ; //editprotect{|| TRUE } ;
		PARENT oUpStat

	@ 2,1 DCSAY 'Status  (A)ctive (R)escue (F)uture (S)old (D)ead' get oRecord:Upstatus valid{|| editupstat(@oRecord)};
			SAYCOLOR 1,BD_LEMONCREME PICTURE '!'    PARENT oUpStat  GETID 'UPSTAT'

	@ 2,90 DCSAY 'Stock Number If Sold' Get oRecord:STOCKNUM  picture '@!' WHEN{ || IIF(oRecord:Upstatus='S',.t. ,.f.)};
		valid{|| DC_GETREFRESH(GETLIST),.T.} SAYCOLOR 9,0 PARENT oUpStat
	@ 3,1 DCSAY 'Dead deal Code (C)redit (B)ought Elsewhere (I)nventory (N)o Response (O)ther (U)psideDown' get oRecord:causedead ;
		when {|| iif(oRecord:Upstatus='D',.t.,.f.)};
		valid{|| iif(oRecord:Causedead $('CBIONU'),.t.,.f.)} ;
		picture '!' SAYCOLOR 9,0   PARENT oUpStat

	@ 4,1 DCSAY '                             OrigDateIn' get oRecord:beback3 EDITPROTECT{|| TRUE }  PARENT oUpStat


	@ 5,1 DCSAY '              Enter Next Follow up Date' GET oRecord:TICKLE  WHEN{|| IIF( oRecord:Upstatus $('AFR'),.t. ,.f. )}   SAYCOLOR 9,0    PARENT oUpStat


	@ 6,1 DCSAY 'Billing Status (I)nProgress (D)elivered' get oRecord:Billstat PICTURE '@!' valid {|| _chkbillstat(oRecord,@getlist) };
		when {|| iif(oRecord:Upstatus='S',.t.,.f.)};
		picture '!' SAYCOLOR 9,0 PARENT oUpStat


	@ 7,1 DCGROUP oUpinfo Caption 'Prospect Demographic' Size 140,10  PARENT oTabpage0


	@ 1,1 DCSAY '      Firstname' Get oRecord:FIRSTNAME  Saycolor 9,0 picture '@!' PARENT oUpInfo
	@ 1,70 DCSAY 'Middle Initial' get oRecord:MiddleInit Saycolor 9,0 picture '@!'                PARENT oUpInfo
	@ 2,1 DCSAY '       Lastname' Get oRecord:LASTNAME  Saycolor 9,0  picture '@!' PARENT oUpInfo
	@ 3,1 DCSAY '         Street' Get oRecord:STREET  Saycolor 9,0  picture '@!'   PARENT oUpInfo
	@ 3,70 DCSAY '       Zipcode' Get  oRecord:ZIPCODE SAYCOLOR 9,0                PARENT oUpInfo ;
		VALID{|| SEEKZIP(@oRecord),_GETSALESTAXPER(@oRecord),;
		DC_GETREFRESH(GETLIST),.T.};
		POPUP{|| BROWGETZIP(@oRecord),DC_GETREFRESH(GETLIST)};
		POPCAPTION 'ZipSearch ' POPWIDTH 100 POPFONT '10.Courier New' POPSTYLE 1

	@ 4,1 DCSAY '      Citystate' Get oRecord:CITYSTATE  Saycolor 9,0   picture '@!' PARENT oUpInfo
	@ 4,61 DCSAY 'State' get oRecord:STATE                                            PARENT oUpInfo
	@ 4,80 DCSAY 'County ' get oRecord:COUNTY                                            PARENT oUpInfo

	@ 5,1 DCSAY '   Last Date In' Get oRecord:DATEIN                                 PARENT oUpInfo

	@ 5,50 DCSAY ' Beback1' Get oRecord:BEBACK1  WHEN{|| oRecord:UPTYPE < 3}   PARENT oUpInfo
	@ 5,80 DCSAY 'Beback2' Get oRecord:BEBACK2  WHEN{|| IIF(!EMPTY(oRecord:BEBACK1),.T. ,.F.)} PARENT oUpInfo

	@ 6,1 DCSAY '          Email' Get oRecord:EMAIL Saycolor 9,0   picture '@!' valid{|| chkemailaddress(@oRecord:EMAIL,.f.)} PARENT oUpInfo
	@ 7,1 DCSAY '           Area' get oRecord:AREA        PARENT oUpInfo   PICTURE '999' SAYCOLOR 9,0
	@ 7,30 DCSAY 'Main Phone#' get oRecord:UPID  SAYCOLOR 9,0   PARENT oUpInfo PICTURE '@R 999-9999'
	@ 7,60 DCSAY 'AltPhone#'  Get oRecord:WORKPHN  Saycolor 9,0 PARENT oUpInfo   PICTURE '@R 999-999-9999-AAAA'
	@ 7,100 DCSAY 'Cell #' get oRecord:CELLPHONE  Saycolor 9,0 PARENT oUpInfo       PICTURE '@R 999-999-9999'
	@ 8,1 DCSAY '     Contact Method [A]ll Ok [H]omephone [C]ell [W]ork [E]mail [D]ONOTCALL!' get oRecord:PREFMETHOD PARENT oUpInfo;
		PICTURE '@!' ;
		VALID {|| IIF( oRecord:PREFMETHOD $('AHCWED'),.T. ,.F. )} ;
		saycolor GRA_CLR_DARKBLUE,GRA_CLR_WHITE
	/// add pushbuttons
	@ 2,115 DCPUSHBUTTONXP CAPTION 'Record Be Back' ;
	 SIZE 20,2 ;
	 ACTION {|| _MAKEBEBACK(@oRecord),DC_GETREFRESH(GETLIST)} ;
	 CONFIG oConfig1  ;
	 BITMAP 7755      ;
	 PARENT oUpinfo

  @ 4,115 DCPUSHBUTTONXP CAPTION 'Send Record to Vin ' ;
	 SIZE 20,2 ;
	 ACTION {|| Savaquote(@cBody,@oRecord),;
	 UPQTUPDT(oRecord,dDelivdate,cDeltime),;
	 _senduptovin(NEWUPS->(RECNO()))};
	 CONFIG oConfig2  ;
	 BITMAP 7762 ;
	 PARENT oUpinfo


	@ 17,1 DCGROUP oVehinfo CAPTION 'Vehicle Interest' size 140,6.5   PARENT oTabpage0

	@ 1,1 DCSAY '                                                    Vehicle Interest' GET oRecord:INTEREST picture '@!'  PARENT oVehInfo;
		SAYCOLOR 9,0;
		VALID{||CHKINTR(@oRecord:INTEREST,@oRecord:FRANCHISE),dc_getrefresh(getlist),.T.}

	@ 2,1 DCSAY '                Source (F)loor Traffic  (I)nternet Lead  (P)hone Up ' get oRecord:ORIGSOURCE VALID{|| IIF(oRecord:ORIGSOURCE $('FIP'),.T. ,.F. )}  ;
		PICTURE '@!' SAYCOLOR 9,0  ;
		GETTOOLTIP('You must enter the source of this Up. This will be permanently stored as the original source.')  PARENT oVehInfo

	@ 3,1 DCSAY 'UP Type 1=NewCust 2=Cust/Refer 3=Beback/1 4=Beback/2. 5=iNet 6=Phone' GET oRecord:UPTYPE PICTURE '9'   PARENT oVehInfo;
		 VALID{|| _CHK4TYPECHANGE(nCurrUptype,@oRecord),DC_GETREFRESH(GETLIST),.t.} ;
		 RANGE 1,6  SAYCOLOR 9,0  ;
		 GETTOOLTIP('For floor traffic always use 123 or 4. For iNet or phone sources always use 5 or 6 UNTIL the up ; actually comes in. THEN change it to a 1 or 2 and the up will be eligible in the up count for the period.')
	@ 4,1 DCSAY '                      Enter N For New Vehicle Or U For Used vehicle.' GET oRecord:NEWUSED PICTURE '@!' VALID{|| CHKNU(oRecord:NEWUSED)} SAYCOLOR 9,0   ;
		PICTURE '@!' PARENT oVehInfo
	@ 5,1 DCSAY '                      Advertising Source or Press Enter for Choices.' GET oRecord:ADSOURCE  PARENT oVehInfo;
		SAYCOLOR 9,0;
		PICTURE '@!' ;
		valid {|| aslk(@oRecord:ADSOURCE),dc_getrefresh(getlist),.T.}

	@ 23.5,1 DCGROUP oProgInfo CAPTION 'Other Prospect Info' size 70,8.3  PARENT oTabpage0


	@ 1,1 DCSAY '      Comments'  Get oRecord:COMMENTS  Saycolor 9,0 PARENT oProgInfo
	@ 2,1 DCSAY ' More Comments' Get oRecord:COMMENT2  Saycolor 9,0 PARENT oProgInfo
	@ 3,1 DCSAY '   T.O.Manager'  Get  oRecord:TOMGR picture '@!' Saycolor 9,0 PARENT oProgInfo;
		VALID{|| IIF( !EMPTY(oRecord:TOMGR) ,;
		oRecord:TOED:=.T. ,oRecord:TOED:=.F. ),DC_GETREFRESH(GETLIST),.T.}

	@ 4,1 DCSAY ' Demo Ride Y/N' Get oRecord:DEMO PARENT oProgInfo Picture '@! L'  Saycolor 9,0 VALID{|| IIF(oRecord:DEMO='Y' ,;
		oRecord:DEMOED:=.T. ,oRecord:DEMOED:=.F. ),DC_GETREFRESH(GETLIST),.T.}
	@ 4,30 DCSAY 'Stk# Demonstrated ' Get oRecord:DEMOSTK  picture '@!' WHEN{|| IIF(oRecord:DEMO='Y' ,.T. ,.F. )};
		SAYCOLOR 9,0 PARENT oProgInfo

	@ 5,1 DCSAY 'If Down Deal Enter Down Code 0,1,2,3,4,5,6' Get oRecord:Downdeal PARENT oProgInfo Saycolor 9,0 Picture '9' Range 0,6
	@ 6,1 DCSAY 'After Delivery Call Made Y/N              ' get oRecord:SENDLTR PICTURE '@! L' Saycolor 9,0;
		 WHEN {|| IIF(oRecord:Billstat='D',.T.,.F. )} PARENT oProgInfo   ;


	@ 23.5,71 DCGROUP oProgInfo2 CAPTION 'Process Progress' size 70,8.3   PARENT oTabpage0



	@ 1,1 DCSAY '         Appointment' get oRecord:APPOINT  EDITPROTECT{|| TRUE }  NOTABSTOP PICTURE 'Y' PARENT oProgInfo2 ;
		  SAYCOLOR{|| IIF(oRecord:APPOINT ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

	/*
	@ 6,24 DCSAY 'Showed Up' get lShowed EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo ;
		  SAYCOLOR {|| _CHK4SHOWCLR(oRecord:UPID)}    */

	@ 2,1 DCSAY '           Demo Ride' get oRecord:DEMOED    EDITPROTECT{|| TRUE }  NOTABSTOP   PICTURE 'Y' PARENT oProgInfo2 ;
		SAYCOLOR{|| IIF(oRecord:DEMOED ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

	@ 3,1 DCSAY '     TOed to Manager' get oRecord:TOED  EDITPROTECT{|| TRUE } NOTABSTOP  PICTURE 'Y' PARENT oProgInfo2;
		SAYCOLOR{|| IIF(oRecord:TOED ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

	@ 4,1 DCSAY '          Written Up'  get oRecord:WRITEUP    EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo2;
		SAYCOLOR{|| IIF(oRecord:WRITEUP ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

	@ 5,1 DCSAY '                Sold'  get oRecord:SOLD            EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo2;
		SAYCOLOR{|| IIF(oRecord:SOLD ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,BD_SALMON} )}

   @ 6,1 DCSAY '           Delivered'  get lDelivered       EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo2 ;
		SAYCOLOR{|| IIF(lDelivered ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,GRA_CLR_PALEGRAY} )}

	@ 7,1 DCSAY '  Post Delivery Call' get oRecord:SENDLTR EDITPROTECT{|| TRUE } NOTABSTOP PICTURE 'Y' PARENT oProgInfo2 ;
		SAYCOLOR{|| IIF(oRecord:SENDLTR='Y' ,{GRA_CLR_BLACK,BD_KEYLIME} ,{GRA_CLR_BLACK,GRA_CLR_PALEGRAY} )}

	@ 4,40 DCPUSHBUTTONXP CAPTION 'Record Appointment ' ;
	 SIZE 20,2 ;
	 ACTION {|| _makeupappt(oRecord)};
	 CONFIG oConfig2  ;
	 BITMAP 8062 ;
	 PARENT oProgInfo2


	@ 32,0 DCTOOLBAR AoToolBar    PARENT oTabpage0                           ;
		SIZE 140,2 BUTTONSIZE 12,2

/*
IF MANAGER()
	DCADDBUTTON CAPTION 'Payment;Calculator'                              ;
		ACTION {|| QUICK(0,0,0,0,0,0,0,0,0,0,.T.,0)  };
		PARENT AoToolBar      ;
		HIDE {|| IF(!MANAGER(),.T.,.F.)}
ENDIF  */

DCADDBUTTONXP CAPTION 'Comments '                              ;
	ACTION {|| UPNOTES()  };
	PARENT AoToolBar   ;
	CONFIG oConfig1

DCADDBUTTONxp CAPTION 'Print; UPSheet '                              ;
	 ACTION {|| Savaquote(@cBody,@oRecord),;
	 UPQTUPDT(oRecord,dDelivdate,cDeltime),;
	PRINTUPSHEET(NEWUPS->(RECNO()))};
	PARENT AoToolBar ;
	CONFIG oConfig1



DCADDBUTTONXP CAPTION 'Buyers ;Order '                              ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		BuyersOrder(NEWUPS->(RECNO()),.f.,@oRecord),DC_GETREFRESH(GETLIST)};
		PARENT AoToolBar    ;
	   CONFIG oConfig1


IF MANAGER()
  DCADDBUTTONXP CAPTION 'Final Buyers; Order '                              ;
		SIZE 15                                              ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		BuyersOrder(NEWUPS->(RECNO()),.t.,@oRecord),DC_GETREFRESH(GETLIST)};
		PARENT AoToolBar ;
	   CONFIG oConfig1

ENDIF

DCADDBUTTONXP CAPTION 'Appraisal; Form '                              ;         /// this will update Newups appraisal record
		ACTION {|| NEWUPS->(DBGOTO(oRecord:Record_Number)),;
	   NEWUPS->(DC_DBGATHER(oRecord)),;
		_appraiseform(NEWUPS->(RECNO()),oRecord:UPID,oRecord:SALESMAN,oRecord,.f.),;
		NEWUPS->(DBGOTO(oRecord:Record_Number)),;
	   NEWUPS->(DC_DBGATHER(oRecord))};
		PARENT AoToolBar ;
	   CONFIG oConfig1


DCADDBUTTONXP CAPTION 'Review; eMails '                              ;
		ACTION {|| BROWEMAILLOG('',NEWUPS->EMAIL)} ;
		PARENT AoToolBar   ;
	    CONFIG oConfig1


DCADDBUTTONXP CAPTION 'Appointment; History '                              ;
	ACTION {||_chkapptsbrowse(NEWUPS->UpId)};
	PARENT AoToolBar ;
	CONFIG oConfig1

DCADDBUTTONXP CAPTION 'InspObj' ;
        PARENT AOTOOLBAR ;
        ACTION {|| dc_arrayview(oRecord)  }   ;
		  CONFIG oConfig1




DCADDBUTTONXP CAPTION 'Exit ;No Save '                              ;
		PARENT AoToolbar                                                                            ;
		ACTION {|| DC_READGUIEVENT(DCGUI_EXIT_ABORT,GETLIST)} ;
		COLOR 1,BD_SALMON  ;
		CONFIG oConfig1

	 DCADDBUTTONXP CAPTION 'Save and Exit '                              ;
		PARENT AoToolbar                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}  ;
		COLOR 1,BD_KEYLIME ;
		CONFIG oConfig1



@ 0,0 DCTABPAGE oTabpage1 CAPTION 'Desking Info ' ;
		RELATIVE oTabpage0 ;
		size 150,34  ;
		GOTFOCUS{|| _REFRESHREBATES(@oRecord),;
		_GETSALESTAXPER(@oRecord),Chkpric(@oRecord,@nProfit,lWarn),Addup(@oRecord),calccod(@oRecord),;
		;//_RecalcAll(oRecord,@nProfit,nLoad),;
		DC_GETREFRESH(GETLIST)}


@ 1.3,1 DCSAY 'Phone#/Upid '  GET oRecord:Upid   EDITPROTECT{||.T.} PARENT oTabpage1
@ 1.3,30 DCSAY 'Area Code ' GET oRecord:Area     EDITPROTECT{||.T.} PARENT oTabpage1
@ 1.3,50 DCCHECKBOX lWarn PROMPT 'Check or Uncheck to ENABLE or DISABLE Gross Warning' ;
	COLOR 1,BD_LINEN   PARENT oTabpage1

@ 2.3,1 DCSAY 'Lastname   ' GET oRecord:Lastname EDITPROTECT{||.T.} PARENT oTabpage1
@ 2.3,70 DCSAY 'Firstname ' GET oRecord:Firstname EDITPROTECT{||.T.} PARENT oTabpage1

/// put up a check box to enable disable warnings


@ 3.5,0 DCGROUP oDealinfo CAPTION  'Sale Details Section' SIZE 140,8.3 PARENT oTabpage1
@ 1,2 DCRADIO oRecord:quotetype VALUE 'R' PROMPT 'Retail' COLOR 1,BD_LINEN ;
      PARENT oDealInfo
@ 1,20 DCRADIO oRecord:quotetype VALUE 'L' PROMPT 'Lease' COLOR 1,BD_LEDGER ;
      PARENT oDealInfo
@ 1,40 DCRADIO oRecord:quotetype VALUE 'C' PROMPT 'Cash Deal' COLOR 0,BD_DARKBLUE ;
		       PARENT oDealInfo

@ 1,70 DCRADIO oRecord:NEWUSED VALUE 'N' PROMPT 'New Vehicle' COLOR 1,BD_PINKSAND ;
		ACTION{|| _SetAutoCharges(@oRecord,'N'),DC_GETREFRESH(GETLIST)};
      PARENT oDealInfo

@ 1,90 DCRADIO oRecord:NEWUSED VALUE 'U' PROMPT 'Used Vehicle' COLOR 1,BD_WHEAT ;
		ACTION{|| _SetAutoCharges(@oRecord,'U'),DC_GETREFRESH(GETLIST)};
      PARENT oDealInfo

@ 1,115 DCPUSHBUTTON PROMPT 'Browse Inventory';
	SIZE 20,2;
	ACTION {|| Vehbrow(oRecord:NEWUSED,@oRecord:STKQUOTED),;
	Chkstknum(@oRecord,@Getlist,.T.),dc_getrefresh(getlist)}  ;
	PARENT oDealInfo


@ 3,1 DCSAY '            Stock Number/LOCATE/ORDER/LBO' Get oRecord:STKQUOTED PARENT oDealinfo PICTURE '@!' Valid{|| Chkstknum(@oRecord,@Getlist,.T.)};
	   GETID 'STOCKNUMBER'
@ 3,70 DCSAY 'Vin' Get oRecord:QUOTEVIN PARENT oDealinfo
//@ 8,90 DCSAY 'Tissue' get oRecord:QUOTETISS picture '999999' editprotect{|| .t.}  NOTABSTOP
@ 3,97 DCSAY 'Mileage' get oRecord:QUOTEMILES PARENT oDealinfo PICTURE '999999'
@ 4,1 DCSAY 'Enter Year' Get oRecord:YRQUOTED PARENT oDealinfo
@ 4,25 DCSAY 'Enter Vehicle Quoted' Get oRecord:VEHQUOTED PARENT oDealinfo picture 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'


/// ADD PUSH BUTTON TO VIEW VEHICLE
@ 5.5,115 DCPUSHBUTTONXP CAPTION 'View Vehicle;Details '                              ;
		SIZE 12,2;
		CARGO 'Cancel' ;
		PARENT oDealInfo                                              ;
		ACTION {|| IIF( oRecord:NEWUSED='U',RUNINNEWTHREAD({|| UCINQFROMUPS(oRecord:STKQUOTED)}) ,RUNINNEWTHREAD({|| NCIINQFROMUPS(oRecord:STKQUOTED)}))}


@ 5,1 DCSAY 'MSRP              ' Get oRecord:QUOTEMSRP PARENT oDealinfo Picture '999999.99'
@ 5,40 DCSAY 'Invoice' get oRecord:QUOTETISS PARENT oDealinfo picture '999999' editprotect{|| .t.}  NOTABSTOP
@ 5,60 DCSAY '     County' get oRecord:COUNTY valid{|| _BROWCOUNTY(@oRecord,@GETLIST)}PARENT oDealinfo
@ 5,97 DCSAY 'Tax%' Get oRecord:TAXPER PARENT oDealinfo Picture '99.999' ;
		 Valid{|| _RecalcAll(@oRecord,nProfit,nLoad),_WarnCtDeal(oRecord),;
	    DC_GETREFRESH(GETLIST),SETAPPFOCUS(DC_GETOBJECT(GETLIST,'CREBDESC1')),.T.}


@ 6,1 DCSAY 'Enter Price Quoted' Get oRecord:QUOTED PARENT oDealinfo Saycolor BD_CORNFLOWERBLUE,0 ;
	Valid{|| _RecalcAll(@oRecord,@nProfit,nLoad),;
	DC_GETREFRESH(GETLIST),.T.} Picture '999999.99'

IF MANAGER()
 @ 6,40 DCSAY '    Net' get nProfit  editprotect {|| .t.}  PICTURE '99999' NOTABSTOP PARENT oDealinfo
ENDIF

@ 6,60 DCSAY 'Tax Amount ' get oRecord:TAX editprotect{|| .t.}  PARENT oDealinfo picture '99999.99' NOTABSTOP


@ 7,1 DCSAY '   Total Cash Down' Get oRecord:DOWNPAY PARENT oDealinfo Picture '99999.99' WHEN{|| IIF( oRecord:QUOTETYPE='R' ,.T.,.F.)};
	Valid{|| Addup(@oRecord),calccod(@oRecord),DC_GETREFRESH(GETLIST),SETAPPFOCUS(DC_GETOBJECT(GETLIST,'TRADE')),.T.};
	NOTABSTOP ;
	Saycolor 1,BD_KEYLIME
@ 7,65 DCSAY 'Enter 2nd SM Id IF this is a split' get oRecord:SM2         PARENT oDealInfo



@ 11.8,0 DCGROUP oTrade CAPTION 'Trade In Information' SIZE 140,4 PARENT oTabpage1
@ 1,1 DCSAY '   Trade Year' Get oRecord:TRADEYR PARENT oTrade GETID 'TRADE';
	POPUP{|| CLEARTRADE(@oRecord),DC_GETREFRESH(GETLIST)};
		POPCAPTION 'ClearInfo' POPWIDTH 60 POPFONT '7.Courier New' POPSTYLE 1

@ 1,30 DCSAY '      Trade Make' get oRecord:TRADEMAKE  PARENT oTrade PICTURE '@!' when{|| IIF( !empty(oRecord:TRADEYR),.t. ,.f. )}
@ 1,70 DCSAY 'Trade Model' Get oRecord:TRADEDESC  PARENT oTrade PICTURE '@!' when{|| IIF( !empty(oRecord:TRADEYR),.t. ,.f. )}
@ 2,1 DCSAY 'Trade Mileage' get oRecord:TRADEMILES PARENT oTrade PICTURE '999999' when{|| IIF( !empty(oRecord:TRADEYR),.t. ,.f. )}
@ 2,30 DCSAY '     Trade Color'    get oRecord:TRADECOLOR  PARENT oTrade PICTURE '@!'  when{|| IIF( !empty(oRecord:TRADEYR),.t. ,.f. )}
@ 2,70 DCSAY '  Trade Vin' get oRecord:TRADEVIN   PARENT oTrade PICTURE '@!' when{|| IIF( !empty(oRecord:TRADEYR),.t. ,.f. )} ;
	 Valid {|| Testvin(oRecord:TRADEVIN,oRecord:TRADEYR)}

@ 1,110 DCPUSHBUTTONXP CAPTION 'Appraisal Form '                              ;
		SIZE 12,2;
		PARENT oTrade                                     ;
		ACTION {|| 	_appraiseform(NEWUPS->(RECNO()),oRecord:UPID,oRecord:SALESMAN,@oRecord,.f.),;
					   Addup(@oRecord),calccod(@oRecord),Chkpric(@oRecord,@nProfit,lWarn),dc_getrefresh(getlist)}    ;
		CONFIG oConfig1


IF MANAGER()
@ 3,1 DCSAY '    Trade ACV' Get oRecord:TRADEACV PARENT oTrade Picture '99999.99' when{|| IIF( !empty(oRecord:TRADEYR),.t. ,.f. )};
  valid{|| Addup(@oRecord),calccod(@oRecord),Chkpric(@oRecord,@nProfit,lWarn),dc_getrefresh(getlist),.t.}
ENDIF
@ 3,30 DCSAY ' Trade Allowance' Get oRecord:TRADEQUOTE PARENT oTrade Picture '99999.99' when{|| IIF( !empty(oRecord:TRADEYR),.t. ,.f. )} ;
	 valid{|| Addup(@oRecord),calccod(@oRecord),Chkpric(@oRecord,@nProfit,lWarn),dc_getrefresh(getlist),.t.}

@ 3,70 DCSAY ' Trade Lien' Get oRecord:LIEN PARENT oTrade Picture '99999.99' when{|| IIF( !empty(oRecord:TRADEYR),.t. ,.f. )};
	valid{|| Addup(@oRecord),calccod(@oRecord),Chkpric(@oRecord,@nProfit,lWarn),dc_getrefresh(getlist),.t.}


@ 16,0 DCGROUP oFinance CAPTION 'Finance Information' size 70,7.5 PARENT oTabpage1
 //@ 1,1 DCSAY 'Interest Rate' get oRecord:INTRATE PARENT oFinance picture '99.99' WHEN{|| IIF( !oRecord:QUOTETYPE='C',.T.,.F.)}

 @ 1,1 DCSAY 'Amount Financed' Get oRecord:FINANCE PARENT oFinance Picture '999999.99' Saycolor 1,BD_KEYLIME Getcolor 1,BD_LEMONCREME ;
		EDITPROTECT{|| TRUE }

 @ 1,31 DCSAY '       Finance Comp' get oRecord:LCOMPANY PARENT oFinance WHEN{|| IIF( !oRecord:QUOTETYPE='C',.T.,.F.)}


 @ 2,31 DCSAY '     Rebates Quoted' Get oRecord:REBATE Picture '99999.99' PARENT oFinance editprotect {|| .t.}  NOTABSTOP

 @ 3,31 DCSAY 'Out the Door C.O.D.' get oRecord:COD PARENT oFinance picture '@R 999999.99' NOTABSTOP EDITPROTECT{|| TRUE }



@ 2,1 DCSAY '  36 Mos Payment' Get  oRecord:PAY36 PARENT oFinance Picture '9999.99' When {|| oRecord:QUOTETYPE ='R'}
@ 3,1 DCSAY '  48 Mos Payment' Get  oRecord:PAY48 PARENT oFinance Picture '9999.99' When {|| oRecord:QUOTETYPE ='R'}
@ 4,1 DCSAY '  60 Mos Payment' Get  oRecord:PAY60 PARENT oFinance Picture '9999.99' When {|| oRecord:QUOTETYPE ='R'}
@ 5,1 DCSAY '  72 Mos Payment' Get  oRecord:PAY72 PARENT oFinance Picture '9999.99' When {|| oRecord:QUOTETYPE ='R'}
@ 6,3 DCGET   oRecord:Termodd EDITPROTECT{|| TRUE }   PARENT oFinance Picture '99'
@ 6,6 DCSAY  'Mos Payment' PARENT oFinance When {|| oRecord:QUOTETYPE ='R'}
@ 6,21 DCGET oRecord:PAYODD When {|| oRecord:QUOTETYPE ='R'} PARENT oFinance picture '9999.99'

@ 16,71 DCGROUP oLeaseSum CAPTION 'Lease Information' size 70,7.5 PARENT oTabpage1


//@ 22,0 DCGROUP oLease CAPTION 'Lease Info' size 140,2  PARENT oTabpage1
@ 1,1 DCSAY '    Lease Pay Quoted' Get oRecord:LPAYMENT2 PARENT oLeasesum Picture '9999.99'  ;
	When {|| oRecord:QUOTETYPE ='L'}

@ 2,1 DCSAY 'Lease Rebates Quoted' Get oRecord:LREBATE Picture '99999.99' PARENT oLeasesum editprotect {|| .t.}  NOTABSTOP

@ 3,1 DCSAY '          Lease Term' Get oRecord:LTERM PARENT oLeasesum Picture '99' ;
	When {|| oRecord:QUOTETYPE ='L'}


@ 4,1 DCSAY '         Lease Miles' Get oRecord:LMILES PARENT oLeasesum Picture '99999' ;
	When {|| oRecord:QUOTETYPE ='L'}

@ 5,1 DCSAY '    Capitalized Cost' Get oRecord:CAPCOST PARENT oLeasesum Picture '999999.99' ;
	When {|| oRecord:QUOTETYPE ='L'}

@ 6,1 DCSAY '  Out The Door C.O.D' Get oRecord:dueatsign PARENT oLeasesum Picture '99999.99' ;
	When {|| oRecord:QUOTETYPE ='L'}

//@ 5,110 DCSAY 'CapFees/Tax ' Get oRecord:CAPFEES PARENT oFinance Picture '@! L' ;
//	When {|| oRecord:QUOTETYPE ='L'}

IF MANAGER()

		@ 23.5,0 DCGROUP oAftsale CAPTION 'Aftersale Charges ' size 140,3.1   PARENT oTabpage1
		DCSETGROUP TO 'AFTSALE'

		@ 1,1 DCGET oRecord:SERVCONT editprotect{|| TRUE }  NOTABSTOP  PARENT oAftsale pict '!!!!!!!!!!!!!'
		@ 2,1 DCGET oRecord:SERVCONAM editprotect{|| TRUE } PARENT oAftsale Picture '99999.99'   NOTABSTOP

		@ 1,15 DCGET oRecord:AFTSALE1 editprotect{|| TRUE }  PARENT oAftsale                pict '!!!!!!!!!!!!!'
		@ 2,15 DCGET oRecord:AFTSALEAM1 editprotect{|| TRUE } PARENT oAftsale Picture '99999.99'

		@ 1,30 DCGET oRecord:AFTSALE2 editprotect{|| TRUE }    NOTABSTOP PARENT oAftsale    pict '!!!!!!!!!!!!!'
		@ 2,30 DCGET oRecord:AFTSALEAM2 editprotect{|| TRUE } PARENT oAftsale Picture '99999.99' NOTABSTOP

		@ 1,45 DCGET oRecord:AFTSALE3 editprotect{|| TRUE } NOTABSTOP PARENT oAftsale       pict '!!!!!!!!!!!!!'
		@ 2,45 DCGET oRecord:AFTSALEAM3 NOTABSTOP EDITPROTECT{|| TRUE } PARENT oAftsale Picture '99999.99'


		@ 1,60 DCGET oRecord:AFTSALE4 NOTABSTOP EDITPROTECT{|| TRUE }   PARENT oAftsale     pict '!!!!!!!!!!!!!'
		@ 2,60 DCGET oRecord:AFTSALEAM4 NOTABSTOP EDITPROTECT{|| TRUE } PARENT oAftsale Picture '99999.99'

		@ 1,75 DCGET oRecord:AFTSALE5 NOTABSTOP EDITPROTECT{|| TRUE }  PARENT oAftsale      pict '!!!!!!!!!!!!!'
		@ 2,75 DCGET oRecord:AFTSALEAM5 NOTABSTOP EDITPROTECT{|| TRUE } PARENT oAftsale Picture '99999.99'

		@ 1,90 DCGET oRecord:AFTSALE6 NOTABSTOP EDITPROTECT{|| TRUE }  PARENT oAftsale      pict '!!!!!!!!!!!!!'
		@ 2,90 DCGET oRecord:AFTSALEAM6 NOTABSTOP EDITPROTECT{|| TRUE } PARENT oAftsale Picture '99999.99'

		@ 1,105 DCSAY 'Addl Tax'                                 PARENT oAftsale
		@ 2,105 DCGET oRecord:AFTSALETAX NOTABSTOP EDITPROTECT{|| TRUE }  Picture '99999.99' PARENT oAftsale

		@ 1,120 DCSAY 'Deposit     '                                                       PARENT oAftsale
		@ 2,120 DCGET oRecord:DEPOSIT picture '@R 999999.99' NOTABSTOP EDITPROTECT{|| TRUE } PARENT oAftsale

		DCSETGROUP TO
	ENDIF
		@ 27,0 DCGROUP oFees CAPTION 'Other Fees/Buyers Order Comments ' size 140,5   PARENT oTabpage1

		DCSETGROUP TO 'FEES'

		@ 1,1 DCSAY 'Inspection  '  PARENT oFees
		@ 2,1 DCGET  oRecord:INSPECTION NOTABSTOP EDITPROTECT{|| TRUE } PARENT oFees Picture '99999.99'

		@ 1,15 DCSAY 'Registration'  PARENT oFees
		@ 2,15 DCGET oRecord:REGFEE  NOTABSTOP EDITPROTECT{|| TRUE }    PARENT oFees  Picture '99999.99'

		@ 1,30 DCSAY 'Processing  '  PARENT oFees
		@ 2,30 DCGET oRecord:PROCFEE NOTABSTOP EDITPROTECT{|| TRUE }    PARENT oFees Picture '99999.99'

		@ 1,45 DCSAY 'Other Tax   '  PARENT oFees
		@ 2,45 DCGET oRecord:WASTEFEE NOTABSTOP EDITPROTECT{|| TRUE }   PARENT oFees Picture '99999.99'

		@ 1,60 DCSAY 'Total '  PARENT oFees
		@ 2,60 DCGET nTotNonVeh  NOTABSTOP EDITPROTECT{|| TRUE } PARENT oFees Picture '99999.99'

		@ 3,60 dcpushbutton caption 'WTF' size 5,1 parent oFees ;
			action {|| _stranalyser(oRecord:bocomment)}

		@ 1,71 DCMULTILINE oRecord:BOCOMMENT SIZE 60,4 FONT '8.Courier New' maxlines 4 LINELENGTH 60 ;
			;// WHEN {|| _reinitstr(@oRecord:bocomment)};
			  NOVERTSCROLL NOHORIZSCROLL PARENT oFees

		DCSETGROUP TO


@ 32,1 DCTOOLBAR oToolbar SIZE 80,1.5 BUTTONSIZE 15,1.5  PARENT oTabpage1

	 DCADDBUTTONXP CAPTION 'Print Quote '                              ;
		PARENT oToolbar                                              ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		PRTQUOTE(oRecord:QUOTEVIN )}    ;
		CONFIG oConfig1


    DCADDBUTTONXP CAPTION 'Buyers Order '                              ;
		PARENT oToolbar                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		BuyersOrder(NEWUPS->(RECNO()),.f.,oRecord)} ;
		CONFIG oConfig1


	DCADDBUTTONXP CAPTION 'Exit ;No Save '                              ;
		PARENT oToolbar                                                                            ;
		ACTION {|| DC_READGUIEVENT(DCGUI_EXIT_ABORT,GETLIST)} ;
		COLOR 1,BD_SALMON ;
		CONFIG oConfig1

	 DCADDBUTTONXP CAPTION 'Save and Exit '                              ;
		PARENT oToolbar                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
      DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}  ;
		COLOR 1,BD_KEYLIME  ;
		CONFIG oConfig1






	DCADDBUTTONXP CAPTION 'Send Record to Vin ' ;
	 SIZE 20,1.5 ;
	 PARENT oToolbar ;
	 ACTION {|| UPQTUPDT(@oRecord,dDelivdate,cDeltime),_senduptovin(NEWUPS->(RECNO()))};
	 BITMAP 8062    ;
	 CONFIG oConfig1


@ 0,0 DCTABPAGE oTabpageCalcs CAPTION 'Side By Side Calcs ';
		RELATIVE oTabpage1 ;
		size 150,34  ;
		GOTFOCUS{|| _REFRESHREBATES(oRecord),;
		_GETSALESTAXPER(@oRecord), ;
		_RecalcAll(@oRecord,@nProfit,nLoad),;
		DC_GETREFRESH(GETLIST)}


@ 1.5,1 DCGROUP oRetailCalcs CAPTION  'Retail Calculations Section' SIZE 70,20 PARENT oTabpageCalcs
@ 1,1 DCSAY '               MSRP' Get oRecord:QUOTEMSRP PARENT oRetailCalcs Picture '999999.99'
@ 2,1 DCSAY '      Price Quoted ' Get oRecord:QUOTED ;
      Valid{|| _RecalcAll(@oRecord,@nProfit,nLoad),;
		DC_GETREFRESH(GETLIST),.T.} Picture '999999.99';
      PARENT   oRetailCalcs

IF MANAGER()
  @ 2,40 DCSAY ' Net' get nProfit  editprotect {|| .t.}  PICTURE '999999.99' NOTABSTOP PARENT oRetailCalcs
  @ 3,1 DCSAY '    Trade Allowance' Get oRecord:TRADEQUOTE ;
	   Valid{|| _RecalcAll(@oRecord,@nProfit,nLoad),;
		DC_GETREFRESH(GETLIST),.T.} Picture '999999.99';
      PARENT oRetailCalcs
  @ 3,40 DCSAY ' ACV' get oRecord:TRADEACV  editprotect {|| .t.}  PICTURE '999999.99' NOTABSTOP PARENT oRetailCalcs  GETCOLOR{|| IIF( oRecord:TRADEACV=0,{1,BD_SALMON} ,{1,BD_KEYLIME} )}
ENDIF
//@ 4,1 DCSAY ' Taxable Amount' Get oRecord:TRADEQUOTE PARENT oRetailCalcs Picture '999999.99'
@ 4,1 DCSAY '          Sales Tax' get oRecord:TAX        PARENT oRetailCalcs Picture '999999.99'
@ 5,1 DCSAY '        DownPayment' get oRecord:DOWNPAY ;
		Valid{|| _RecalcAll(@oRecord,@nProfit,nLoad),;
		DC_GETREFRESH(GETLIST),.T.} Picture '999999.99';
	   PARENT oRetailCalcs Picture '999999.99'

@ 6,1 DCSAY '            Rebates' get oRecord:REBATE     PARENT oRetailCalcs Picture '999999.99'  EDITPROTECT{|| TRUE }

@ 7,1 DCSAY '         LienPayoff' get oRecord:LIEN ;
		Valid{|| _RecalcAll(@oRecord,@nProfit,nLoad),;
		DC_GETREFRESH(GETLIST),.T.} Picture '999999.99';
		PARENT oRetailCalcs Picture '999999.99'

@ 8,1 DCSAY '    Cash Amount Due' get oRecord:COD     PARENT oRetailCalcs Picture '999999.99'
@ 9,1 DCSAY '    Financed Amount' GET oRecord:FINANCE PARENT oRetailCalcs Picture '999999.99'


@ 11,1 DCSAY 'Pmt 36' Get  oRecord:PAY36 PARENT oRetailCalcs Picture '99999.99'
@ 11,25 DCSAY 'Interest Rate '  get oRecord:Rate36 EDITPROTECT{|| TRUE } SAYCOLOR {|| IIF( oRecord:Rate36=0,{1,BD_SALMON} ,{1,BD_KEYLIME} )} PARENT oRetailCalcs

@ 12,1 DCSAY 'Pmt 48' Get  oRecord:PAY48 PARENT oRetailCalcs Picture '99999.99'
@ 12,25 DCSAY 'Interest Rate '  get oRecord:Rate48 EDITPROTECT{|| TRUE } SAYCOLOR {|| IIF( oRecord:Rate48=0,{1,BD_SALMON} ,{1,BD_KEYLIME} )} PARENT oRetailCalcs

@ 13,1 DCSAY 'Pmt 60' Get  oRecord:PAY60 PARENT oRetailCalcs Picture '99999.99'
@ 13,25 DCSAY 'Interest Rate '  get oRecord:Rate60 EDITPROTECT{|| TRUE } SAYCOLOR {|| IIF( oRecord:Rate60=0,{1,BD_SALMON} ,{1,BD_KEYLIME} )} PARENT oRetailCalcs

@ 14,1 DCSAY 'Pmt 72' Get  oRecord:PAY72 PARENT oRetailCalcs Picture '99999.99'
@ 14,25 DCSAY 'Interest Rate '  get oRecord:Rate72 EDITPROTECT{|| TRUE } SAYCOLOR {|| IIF( oRecord:Rate72=0,{1,BD_SALMON} ,{1,BD_KEYLIME} )} PARENT oRetailCalcs

@ 15,1 DCSAY 'Pmt'  get oRecord:TERMODD PICTURE '99'    PARENT oRetailCalcs
@ 15,9 DCGET oRecord:PAYODD  Picture '99999.99' PARENT oRetailCalcs
@ 15,25 DCSAY 'Interest Rate '  get oRecord:INTRATE EDITPROTECT{|| TRUE } SAYCOLOR {|| IIF( oRecord:INTRATE=0,{1,BD_SALMON} ,{1,BD_KEYLIME} )} PARENT oRetailCalcs



@ 1.5,71 DCGROUP oLeaseCalcs CAPTION  'Lease Calculations Section' SIZE 70,20 PARENT oTabpageCalcs


@ 1,1 DCSAY '               MSRP' Get oRecord:QUOTEMSRP PARENT oLeaseCalcs Picture '999999.99'
@ 2,1 DCSAY '      Price Quoted ' Get oRecord:QUOTED ;
		Valid{|| _RecalcAll(@oRecord,@nProfit,nLoad),;
		DC_GETREFRESH(GETLIST),.T.} Picture '999999.99';
      PARENT   oLeaseCalcs

@ 2,40 DCSAY ' Net' get nProfit  editprotect {|| .t.}  PICTURE '999999.99' NOTABSTOP PARENT oLeaseCalcs
@ 3,1 DCSAY '    Trade Allowance' Get oRecord:TRADEQUOTE ;
		Valid{|| _RecalcAll(@oRecord,@nProfit,nLoad),;
		DC_GETREFRESH(GETLIST),.T.} Picture '999999.99';
      PARENT oLeaseCalcs
//@ 4,1 DCSAY ' Taxable Amount' Get oRecord:TRADEQUOTE PARENT oRetailCalcs Picture '999999.99'
//@ 4,1 DCSAY '          Sales Tax' get oRecord:TAX        PARENT oRetailCalcs Picture '999999.99'
@ 5,1 DCSAY '        DownPayment' get oRecord:DOWNPAY  ;        /// allow free form entry
		Valid{|| _RecalcAll(@oRecord,@nProfit,nLoad),;
		DC_GETREFRESH(GETLIST),.T.} Picture '999999.99';
		PARENT oLeaseCalcs Picture '999999.99'

@ 6,1 DCSAY '            Rebates' get oRecord:LREBATE     PARENT oLeaseCalcs Picture '999999.99'  EDITPROTECT{|| TRUE }

@ 7,1 DCSAY '         LienPayoff' get oRecord:LIEN        PARENT oLeaseCalcs Picture '999999.99';
		Valid{|| _RecalcAll(@oRecord,nProfit,nLoad),;
		DC_GETREFRESH(GETLIST),.T.} Picture '999999.99';


@ 8,1 DCSAY '           Cap Cost' get oRecord:CAPCOST     PARENT oLeaseCalcs Picture '999999.99'  EDITPROTECT{|| TRUE }
@ 10,1 DCSAY '        Lease Term ' get oRecord:LTERM       PARENT oLeaseCalcs Picture '999'        EDITPROTECT{|| TRUE }
@ 11,0 DCSAY '       Lease Payment' get oRecord:LPAYMENT2    PARENT oLeaseCalcs Picture '9999.99'    EDITPROTECT{|| TRUE }

@ 32,1 DCTOOLBAR oToolCalc SIZE 75,1.5 BUTTONSIZE 15,1.5  PARENT oTabpageCalcs

	 DCADDBUTTONXP CAPTION 'Print Quote '                              ;
		PARENT oToolCalc           ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
      PRTQUOTE(oRecord:QUOTEVIN )};
		CONFIG oConfig1

    DCADDBUTTONXP CAPTION 'Buyers Order '                              ;
		PARENT oToolCalc                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		BuyersOrder(NEWUPS->(RECNO()),.F.,@oRecord)} ;
		CONFIG oConfig1

  IF manager()
	DCADDBUTTONXP CAPTION 'Final ;Buyers Order '                              ;
		PARENT oToolCalc                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		BuyersOrder(NEWUPS->(RECNO()),.t.,@oRecord)}     ;
		CONFIG oConfig2

  endif

	DCADDBUTTONXP CAPTION 'Exit;No Save '                              ;
		PARENT oToolCalc                                                                            ;
		ACTION {|| DC_READGUIEVENT(DCGUI_EXIT_ABORT,GETLIST)} ;
		COLOR 1,BD_SALMON ;
		CONFIG oConfig1

	DCADDBUTTONXP CAPTION 'Save and Exit '                              ;
		PARENT oToolCalc                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
      DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}  ;
		COLOR 1,BD_KEYLIME  ;
		CONFIG oConfig1


@ 0,0 DCTABPAGE oTabpage2 CAPTION 'Rebates ';
		RELATIVE oTabpageCalcs ;
		size 150,34

@ 3.5,0 DCGROUP oRebate CAPTION  'Retail Rebates Section' SIZE 149,7.5 PARENT oTabpage2

@ 1,1 DCSAY 'FactCode' get oRecord:REBCODE1 PARENT oRebate  PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:REBCODE1,@oRecord:REBATE1,@oRecord:REBDESC1),;
	DC_GETREFRESH(GETLIST),.T.} ;
	POPUP{|| XBROWREBATES(@oRecord:rebcode1,@oRecord:REBDESC1),DC_GETREFRESH(GETLIST)}  ;
		POPCAPTION '??' POPWIDTH 30 POPFONT '8.Courier New' POPSTYLE 1

@ 2,1 DCSAY '  Amount' get oRecord:REBATE1 picture '9999.99' PARENT oRebate ; //when{|| IIF( !empty(oRecord:REBDESC1),.t. ,.f. )};
	valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),;
	calccod(@oRecord),;
	dc_getrefresh(getlist),.t.}
@ 3,1 DCSAY ' Rebate1' get oRecord:REBDESC1 GETID 'CREBDESC1' PARENT oRebate LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}
@ 4,1 DCCHECKBOX lApptols1 PROMPT 'OK Lease'  LOSTFOCUS {|| APPTOLS1(oRecord,lApptols1),;
	oRecord:LREBATE:=oRecord:LREBATE1+oRecord:LREBATE2+oRecord:LREBATE3+oRecord:LREBATE4;
	+oRecord:LREBATE5+oRecord:LREBATE6,Addup(@oRecord),calccod(@oRecord),;
	DC_GETREFRESH(GETLIST)}  PARENT oRebate

@ 1,25 DCSAY 'FactCode' get oRecord:REBCODE2 PARENT oRebate PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:REBCODE2,@oRecord:REBATE2,@oRecord:REBDESC2),;
	DC_GETREFRESH(GETLIST),.T.} ;
	POPUP{|| XBROWREBATES(@oRecord:rebcode2,@oRecord:REBDESC2),DC_GETREFRESH(GETLIST)}  ;
		POPCAPTION '??' POPWIDTH 30 POPFONT '8.Courier New' POPSTYLE 1

@ 2,25 DCSAY '  Amount'  get oRecord:REBATE2 PARENT oRebate picture '9999.99' ;//when{|| IIF( !empty(oRecord:REBDESC2),.t. ,.f. )};
	valid{|| _REFRESHREBATES(@oRecord),;
	Addup(@oRecord),;
	calccod(@oRecord),;
	dc_getrefresh(getlist),.t.}

@ 3,25 DCSAY ' Rebate2' get oRecord:REBDESC2 PARENT oRebate    PARENT oRebate LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}
@ 4,25 DCCHECKBOX lApptols2 PROMPT 'OK Lease' LOSTFOCUS {|| APPTOLS2(oRecord,lApptols2),;
	oRecord:LREBATE:=oRecord:LREBATE1+oRecord:LREBATE2+oRecord:LREBATE3+oRecord:LREBATE4;
	+oRecord:LREBATE5+oRecord:LREBATE6,Addup(@oRecord),calccod(@oRecord),;
   DC_GETREFRESH(GETLIST)}  PARENT oRebate

@ 1,50 DCSAY 'FactCode' get oRecord:REBCODE3 PARENT oRebate PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:REBCODE3,@oRecord:REBATE3,@oRecord:REBDESC3),;
	DC_GETREFRESH(GETLIST),.T.} ;
	POPUP{|| XBROWREBATES(@oRecord:rebcode3,@oRecord:REBDESC3),DC_GETREFRESH(GETLIST)}  ;
		POPCAPTION '??' POPWIDTH 30 POPFONT '8.Courier New' POPSTYLE 1

@ 2,50 DCSAY '  Amount'  get oRecord:REBATE3 PARENT oRebate picture '9999.99'; //when{|| IIF( !empty(oRecord:REBDESC3),.t. ,.f. )};
	valid{|| _REFRESHREBATES(@oRecord),;
	Addup(@oRecord),;
	calccod(@oRecord),;
	dc_getrefresh(getlist),.t.}

@ 3,50 DCSAY ' Rebate3' get oRecord:REBDESC3   PARENT oRebate LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}
@ 4,50 DCCHECKBOX lApptols3 PROMPT 'OK Lease' LOSTFOCUS {|| APPTOLS3(oRecord,lApptols3),;
	oRecord:LREBATE:=oRecord:LREBATE1+oRecord:LREBATE2+oRecord:LREBATE3+oRecord:LREBATE4;
	+oRecord:LREBATE5+oRecord:LREBATE6,Addup(@oRecord),calccod(@oRecord),;
   DC_GETREFRESH(GETLIST)} PARENT oRebate

@ 1,75 DCSAY 'FactCode' get oRecord:REBCODE4 PARENT oRebate PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:REBCODE4,@oRecord:REBATE4,@oRecord:REBDESC4),;
	DC_GETREFRESH(GETLIST),.T.}       ;
	POPUP{|| XBROWREBATES(@oRecord:rebcode4,@oRecord:REBDESC4),DC_GETREFRESH(GETLIST)}  ;
		POPCAPTION '??' POPWIDTH 30 POPFONT '8.Courier New' POPSTYLE 1

@ 2,75 DCSAY '  Amount' get oRecord:REBATE4 PARENT oRebate picture '9999.99'; //when{|| IIF( !empty(oRecord:REBDESC4),.t. ,.f. )};
	valid{|| _REFRESHREBATES(@oRecord),;
	Addup(@oRecord),;
	calccod(@oRecord),;
	dc_getrefresh(getlist),.t.}

@ 3,75 DCSAY ' Rebate4' get oRecord:REBDESC4  PARENT oRebate  LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}
@ 4,75 DCCHECKBOX lApptols4 PROMPT 'OK Lease' LOSTFOCUS {|| APPTOLS4(oRecord,lApptols4),;
	oRecord:LREBATE:=oRecord:LREBATE1+oRecord:LREBATE2+oRecord:LREBATE3+oRecord:LREBATE4;
	+oRecord:LREBATE5+oRecord:LREBATE6,Addup(@oRecord),calccod(@oRecord),;
   DC_GETREFRESH(GETLIST)} PARENT oRebate

@ 1,100 DCSAY ' FactCode' get oRecord:REBCODE5  PARENT oRebate PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:REBCODE5,@oRecord:REBATE5,@oRecord:REBDESC5),;
	DC_GETREFRESH(GETLIST),.T.}          ;
	POPUP{|| XBROWREBATES(@oRecord:rebcode5,@oRecord:REBDESC5),DC_GETREFRESH(GETLIST)}  ;
		POPCAPTION '??' POPWIDTH 30 POPFONT '8.Courier New' POPSTYLE 1

@ 2,100 DCSAY '   Amount' get  oRecord:REBATE5 PARENT oRebate picture '9999.99'; // when{|| IIF( !empty(oRecord:REBDESC5),.t. ,.f. )};
	valid{|| _REFRESHREBATES(@oRecord),;
	Addup(@oRecord),;
	calccod(@oRecord),;
	dc_getrefresh(getlist),.t.}

@ 3,100 DCSAY 'Rebate 5' get oRecord:REBDESC5 PARENT oRebate LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}
@ 4,100 DCCHECKBOX lApptols5 PROMPT 'OK Lease' LOSTFOCUS {|| APPTOLS5(oRecord,lApptols5),;
	oRecord:LREBATE:=oRecord:LREBATE1+oRecord:LREBATE2+oRecord:LREBATE3+oRecord:LREBATE4;
	+oRecord:LREBATE5+oRecord:LREBATE6,Addup(@oRecord),calccod(@oRecord),;
   DC_GETREFRESH(GETLIST)} PARENT oRebate

@ 1,125 DCSAY ' FactCode' get oRecord:REBCODE6  PARENT oRebate PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:REBCODE5,@oRecord:REBATE5,@oRecord:REBDESC5),;
	DC_GETREFRESH(GETLIST),.T.}  ;
	POPUP{|| XBROWREBATES(@oRecord:rebcode6,@oRecord:REBDESC6),DC_GETREFRESH(GETLIST)}  ;
		POPCAPTION '??' POPWIDTH 30 POPFONT '8.Courier New' POPSTYLE 1

@ 2,125 DCSAY '   Amount' get  oRecord:REBATE6 PARENT oRebate picture '9999.99'; // when{|| IIF( !empty(oRecord:REBDESC5),.t. ,.f. )};
	valid{|| _REFRESHREBATES(@oRecord),;
	Addup(@oRecord),;
	calccod(@oRecord),;
	dc_getrefresh(getlist),.t.}

@ 3,125 DCSAY 'Rebate 6' get oRecord:REBDESC6 PARENT oRebate LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}
@ 4,125 DCCHECKBOX lApptols5 PROMPT 'OK Lease' LOSTFOCUS {|| APPTOLS6(oRecord,lApptols6),;
	oRecord:LREBATE:=oRecord:LREBATE1+oRecord:LREBATE2+oRecord:LREBATE3+oRecord:LREBATE4;
	+oRecord:LREBATE5+oRecord:LREBATE6,Addup(@oRecord),calccod(@oRecord),;
   DC_GETREFRESH(GETLIST)} PARENT oRebate



@ 6,1 DCSAY 'Total Retail Rebates' GET oRecord:REBATE PARENT oRebate PICTURE '99999.99'



@ 13.5,0 DCGROUP oLRebate CAPTION  'Lease Rebate Section' SIZE 149,8 PARENT oTabpage2

@ 1,1 DCSAY 'FactCode' get oRecord:LREBCODE1 PARENT oLRebate PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:LREBCODE1,@oRecord:LREBATE1,@oRecord:LREBDESC1),;
	DC_GETREFRESH(GETLIST),.T.} ;
	POPUP{|| XBROWREBATES(@oRecord:lrebcode1,@oRecord:lREBDESC1),DC_GETREFRESH(GETLIST)}  ;
		POPCAPTION '??' POPWIDTH 30 POPFONT '8.Courier New' POPSTYLE 1

@ 2,1 DCSAY '  Amount' get oRecord:LREBATE1 picture '9999.99' PARENT oLRebate ;//when{|| IIF( !empty(oRecord:LREBDESC1),.t. ,.f. )};
	valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
@ 3,1 DCSAY ' Rebate1' get oRecord:LREBDESC1 GETID 'CREBDESC1' PARENT oLRebate  LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}


@ 1,25 DCSAY 'FactCode' get oRecord:LREBCODE2 PARENT oLRebate PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:LREBCODE2,@oRecord:LREBATE2,@oRecord:LREBDESC2),;
	DC_GETREFRESH(GETLIST),.T.} ;
	POPUP{|| XBROWREBATES(@oRecord:lrebcode2,@oRecord:LREBDESC2),DC_GETREFRESH(GETLIST)}  ;
		POPCAPTION '??' POPWIDTH 30 POPFONT '8.Courier New' POPSTYLE 1

@ 2,25 DCSAY '  Amount'  get oRecord:LREBATE2 PARENT oLRebate picture '9999.99'; // when{|| IIF( !empty(oRecord:REBDESC2),.t. ,.f. )};
	valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
@ 3,25 DCSAY ' Rebate2' get oRecord:LREBDESC2 PARENT oLRebate LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}




@ 1,50 DCSAY 'FactCode' get oRecord:LREBCODE3 PARENT oLRebate PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:LREBCODE3,@oRecord:LREBATE3,@oRecord:LREBDESC3),;
	DC_GETREFRESH(GETLIST),.T.}  ;
	POPUP{|| XBROWREBATES(@oRecord:Lrebcode3,@oRecord:LREBDESC3),DC_GETREFRESH(GETLIST)}  ;
		POPCAPTION '??' POPWIDTH 30 POPFONT '8.Courier New' POPSTYLE 1

@ 2,50 DCSAY '  Amount'  get oRecord:LREBATE3 PARENT oLRebate picture '9999.99'; // when{|| IIF( !empty(oRecord:LREBDESC3),.t. ,.f. )};
	valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
@ 3,50 DCSAY ' Rebate3' get oRecord:LREBDESC3   PARENT oLRebate  LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}


@ 1,75 DCSAY 'FactCode' get oRecord:LREBCODE4 PARENT oLRebate PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:LREBCODE4,@oRecord:LREBATE4,@oRecord:LREBDESC4),;
	DC_GETREFRESH(GETLIST),.T.} ;
	POPUP{|| XBROWREBATES(@oRecord:Lrebcode4,@oRecord:LREBDESC4),DC_GETREFRESH(GETLIST)}  ;
		POPCAPTION '??' POPWIDTH 30 POPFONT '8.Courier New' POPSTYLE 1

@ 2,75 DCSAY '  Amount' get oRecord:LREBATE4 PARENT oLRebate picture '9999.99'; //when{|| IIF( !empty(oRecord:LREBDESC4),.t. ,.f. )};
	valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
@ 3,75 DCSAY ' Rebate4' get oRecord:LREBDESC4  PARENT oLRebate   LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}


@ 1,100 DCSAY 'FactCode' get oRecord:LREBCODE5 PARENT oLRebate PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:LREBCODE5,@oRecord:LREBATE5,@oRecord:LREBDESC5),;
	DC_GETREFRESH(GETLIST),.T.};
	POPUP{|| XBROWREBATES(@oRecord:Lrebcode5,@oRecord:LREBDESC5),DC_GETREFRESH(GETLIST)}  ;
		POPCAPTION '??' POPWIDTH 30 POPFONT '8.Courier New' POPSTYLE 1

@ 2,100 DCSAY '  Amount' get  oRecord:LREBATE5 PARENT oLRebate picture '9999.99'; //when{|| IIF( !empty(oRecord:LREBDESC5),.t. ,.f. )};
	valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
@ 3,100 DCSAY ' Rebate5' get oRecord:LREBDESC5 PARENT oLRebate   LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}

@ 1,125 DCSAY 'FactCode' get oRecord:LREBCODE6 PARENT oLRebate PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:LREBCODE6,@oRecord:LREBATE6,@oRecord:LREBDESC6),;
	DC_GETREFRESH(GETLIST),.T.} ;
	POPUP{|| XBROWREBATES(@oRecord:Lrebcode6,@oRecord:LREBDESC6),DC_GETREFRESH(GETLIST)}  ;
		POPCAPTION '??' POPWIDTH 30 POPFONT '8.Courier New' POPSTYLE 1

@ 2,125 DCSAY '  Amount' get  oRecord:LREBATE6 PARENT oLRebate picture '9999.99'; //when{|| IIF( !empty(oRecord:LREBDESC5),.t. ,.f. )};
	valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
@ 3,125 DCSAY ' Rebate6' get oRecord:LREBDESC6 PARENT oLRebate   LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}



@ 5,1 DCSAY 'Total Lease Rebates' GET oRecord:LREBATE PARENT oLRebate PICTURE '99999.99'




@ 32,1 DCTOOLBAR oToolbar2 SIZE 75,1.5 BUTTONSIZE 15,1.5  PARENT oTabpage2

	 DCADDBUTTONXP CAPTION 'Print Quote '                              ;
		PARENT oToolbar2                                              ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
      PRTQUOTE(oRecord:QUOTEVIN )} ;
		CONFIG oConfig1

    DCADDBUTTONXP CAPTION 'Buyers Order '                              ;
		PARENT oToolbar2                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		BuyersOrder(NEWUPS->(RECNO()),.f.,@oRecord)}   ;
		CONFIG oConfig1


	DCADDBUTTONXP CAPTION 'Exit;No Save '                              ;
		PARENT oToolbar2                                                                            ;
		ACTION {|| DC_READGUIEVENT(DCGUI_EXIT_ABORT,GETLIST)} ;
		COLOR 1,BD_SALMON  ;
		CONFIG oConfig1

	DCADDBUTTONXP CAPTION 'Save and Exit '                              ;
		PARENT oToolbar2                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
      DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}  ;
		COLOR 1,BD_KEYLIME  ;
		CONFIG oConfig1



@ 0,0 DCTABPAGE oTabpage3 CAPTION 'Finance Payments ';
		RELATIVE oTabpage2 ;
		size 150,34;
		GOTFOCUS{|| IIF( oRecord:termodd=0 ,oRecord:termodd:=72 ,nil ),;
	   _RecalcAll(oRecord,@nProfit,nLoad),;
		DC_GETREFRESH(GETLIST)	}   /// set odd term to 6 years

	@ 2,15 DCSAY 'Quick Payment Calculator'   PARENT oTabpage3
	@ 5,5 DCSAY '                Enter Financed Amount' Get oRecord:FINANCE Picture "999999" Valid{||Chkfinamt(oRecord:FINANCE)}  Saycolor 9,0 Getcolor 10,0 Getfont '12.Courier'    PARENT oTabpage3
	@ 7,5 DCSAY '                Default Interest Rate' Get oRecord:INTRATE Valid{|| Chkint(oRecord:INTRATE,@oRecord:Rate24,@oRecord:Rate36,@oRecord:Rate48,@oRecord:Rate60)} Picture '999.999';
		PARENT oTabpage3 Saycolor 9,0 Getcolor 10,0 Getfont '12.Courier'
	@ 8,5 DCSAY '       Enter Non Standard Term If Any' Get oRecord:Termodd Picture '999' Saycolor 9,0 Getcolor 10,0 Getfont '12.Courier'  PARENT oTabpage3
	@ 9,5 DCSAY 'Load Financed by Specified Amount Y/N' GET clLoadPayment picture '@! L'                                   PARENT oTabpage3
	@ 10,5 DCSAY '     Amount to Add to Financed Amount' get nLoad picture '9999'  valid{|| iif(clLoadPayment='Y',pf:=oRecord:FINANCE+nload,pf:=oRecord:FINANCE),;		 /// add load if there is one
	calcpayment(60,oRecord:Rate60,@oRecord:Pay60,oRecord:FINANCE+nload),;
	calcpayment(48,oRecord:Rate48,@oRecord:Pay48,oRecord:FINANCE+nload),;
	calcpayment(36,oRecord:Rate36,@oRecord:Pay36,oRecord:FINANCE+nload),;
	calcpayment(72,oRecord:Rate72,@oRecord:Pay72,oRecord:FINANCE+nload),;
	calcpayment(oRecord:termodd,oRecord:IntRate,@oRecord:payodd,oRecord:FINANCE+nload),;
	dc_getrefresh(getlist),.t.}  PARENT oTabpage3

	//@ 11,5 DCSAY 'Rate 24' get oRecord:Rate24  VALID {|| calcpayment(24,oRecord:Rate24,@oRecord:Pay24,oRecord:FINANCE+nload),dc_getrefresh(getlist),.t.} PARENT oTabpage3
	//@ 11,25 DCSAY '24 Months'  get oRecord:Pay24  picture '99999.99' NOTABSTOP EDITPROTECT{|| .t.}                      PARENT oTabpage3

	@ 11,5 DCSAY 'Rate 36' Get oRecord:rate36 VALID {|| calcpayment(36,oRecord:Rate36,@oRecord:Pay36,oRecord:FINANCE+nload),dc_getrefresh(getlist),.t.} PARENT oTabpage3;
	 Picture '999.999'
	@ 11,25 DCSAY '36 Months'  get oRecord:Pay36  picture '99999.99' NOTABSTOP EDITPROTECT{|| .t.}                      PARENT oTabpage3

	@ 13,5 DCSAY 'Rate 48' get oRecord:Rate48 VALID {|| calcpayment(48,oRecord:Rate48,@oRecord:Pay48,oRecord:FINANCE+nload),dc_getrefresh(getlist),.t.} PARENT oTabpage3 ;
		 Picture '999.999'
	@ 13,25 DCSAY '48 Months'  get oRecord:Pay48  picture '99999.99' NOTABSTOP EDITPROTECT{|| .t.}                      PARENT oTabpage3

	@ 15,5 DCSAY 'Rate 60' get oRecord:Rate60 VALID {|| calcpayment(60,oRecord:Rate60,@oRecord:Pay60,oRecord:FINANCE+nload),dc_getrefresh(getlist),.t.} PARENT oTabpage3  ;
		 Picture '999.999'
	@ 15,25  DCSAY '60 Months' get oRecord:Pay60  picture '99999.99' NOTABSTOP EDITPROTECT{|| .t.}                      PARENT oTabpage3

	@ 17,5 DCSAY 'Rate 72' get oRecord:Rate72 VALID {|| calcpayment(72,oRecord:Rate72,@oRecord:Pay72,oRecord:FINANCE+nload),dc_getrefresh(getlist),.t.} PARENT oTabpage3   ;
		 Picture '999.999'
	@ 17,25  DCSAY '72 Months' get oRecord:Pay72  picture '99999.99' NOTABSTOP EDITPROTECT{|| .t.}                      PARENT oTabpage3


	@ 19,5  DCSAY 'RateOdd' PARENT oTabpage3   //get oRecord:Termodd PICTURE '99'  NOTABSTOP                      PARENT oTabpage3
	@ 19,15 DCGET oRecord:intrate picture '999.999'  VALID {|| calcpayment(oRecord:Termodd,oRecord:intrate,@oRecord:Payodd,oRecord:FINANCE+nload),dc_getrefresh(getlist),.t.} PARENT oTabpage3
	@ 19,25 DCGET oRecord:Termodd PARENT oTabpage3 picture '99'
	@ 19,28 DCSAY 'Months' PARENT oTabpage3
	@ 19,37 DCGET oRecord:payodd  picture '99999.99' NOTABSTOP EDITPROTECT{|| .t.}                                     PARENT oTabpage3

	@ 21.5,30 DCGROUP oRebateY CAPTION  'Retail Rebate Section' SIZE 100,9.5 PARENT oTabpage3

		@ 1,1 DCSAY 'FactCode' get oRecord:REBCODE1 PARENT oRebateY  PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:REBCODE1,@oRecord:REBATE1,@oRecord:REBDESC1),;
			DC_GETREFRESH(GETLIST),.T.}
		@ 2,1 DCSAY '  Amount' get oRecord:REBATE1 picture '9999.99' PARENT oRebateY ; //when{|| IIF( !empty(oRecord:REBDESC1),.t. ,.f. )};
			valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
		@ 3,1 DCSAY ' Rebate1' get oRecord:REBDESC1 GETID 'CREBDESC1' PARENT oRebateY LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}

		@ 1,25 DCSAY 'FactCode' get oRecord:REBCODE2 PARENT oRebateY PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:REBCODE2,@oRecord:REBATE2,@oRecord:REBDESC2),;
			DC_GETREFRESH(GETLIST),.T.}
		@ 2,25 DCSAY '  Amount'  get oRecord:REBATE2 PARENT oRebateY picture '9999.99' ;//when{|| IIF( !empty(oRecord:REBDESC2),.t. ,.f. )};
			valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
		@ 3,25 DCSAY ' Rebate2' get oRecord:REBDESC2 PARENT oRebateY    LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}

		@ 1,50 DCSAY 'FactCode' get oRecord:REBCODE3 PARENT oRebateY PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:REBCODE3,@oRecord:REBATE3,@oRecord:REBDESC3),;
			DC_GETREFRESH(GETLIST),.T.}
		@ 2,50 DCSAY '  Amount'  get oRecord:REBATE3 PARENT oRebateY picture '9999.99'; //when{|| IIF( !empty(oRecord:REBDESC3),.t. ,.f. )};
			valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
		@ 3,50 DCSAY ' Rebate3' get oRecord:REBDESC3   PARENT oRebateY LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}

		@ 1,75 DCSAY 'FactCode' get oRecord:REBCODE4 PARENT oRebateY PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:REBCODE4,@oRecord:REBATE4,@oRecord:REBDESC4),;
			DC_GETREFRESH(GETLIST),.T.}
		@ 2,75 DCSAY '  Amount' get oRecord:REBATE4 PARENT oRebateY picture '9999.99'; //when{|| IIF( !empty(oRecord:REBDESC4),.t. ,.f. )};
			valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
		@ 3,75 DCSAY ' Rebate4' get oRecord:REBDESC4  PARENT oRebateY  LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}

		@ 5,1 DCSAY ' FactCode' get oRecord:REBCODE5  PARENT oRebateY PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:REBCODE5,@oRecord:REBATE5,@oRecord:REBDESC5),;
			DC_GETREFRESH(GETLIST),.T.}
		@ 6,1 DCSAY '   Amount' get  oRecord:REBATE5 PARENT oRebateY picture '9999.99'; // when{|| IIF( !empty(oRecord:REBDESC5),.t. ,.f. )};
			valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
		@ 7,1 DCSAY 'Rebate 5' get oRecord:REBDESC5 PARENT oRebateY LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}

		@ 5,25 DCSAY ' FactCode' get oRecord:REBCODE6  PARENT oRebateY PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:REBCODE6,@oRecord:REBATE6,@oRecord:REBDESC6),;
			DC_GETREFRESH(GETLIST),.T.}
		@ 6,25 DCSAY '   Amount' get  oRecord:REBATE6 PARENT oRebateY picture '9999.99'; // when{|| IIF( !empty(oRecord:REBDESC5),.t. ,.f. )};
			valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
		@ 7,25 DCSAY 'Rebate 6' get oRecord:REBDESC6 PARENT oRebateY LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}



		@ 8.5,1 DCSAY 'Total Retail Rebates' GET oRecord:REBATE PARENT oRebateY PICTURE '99999.99'




		@ 11,65 DCPUSHBUTTON CAPTION 'Get To;Payment 36' PARENT oTabpage3;
			SIZE 12,2;
			CARGO 'CANCEL' ;
			ACTION {|| upgettopay(36,oRecord:IntRate,@oRecord:FINANCE,oRecord:Pay36,@oRecord:DownPay,@oRecord:Quoted,oRecord:Taxper,@oRecord),;
			_RecalcAll(@oRecord,@nProfit,nLoad), ;
			dc_getrefresh(getlist)}

		@ 13,65 DCPUSHBUTTON CAPTION 'Get To;Payment 48' PARENT oTabpage3;
			SIZE 12,2;
			CARGO 'CANCEL' ;
			ACTION {|| upgettopay(48,oRecord:IntRate,@oRecord:FINANCE,oRecord:Pay48,oRecord:DownPay,@oRecord:Quoted,oRecord:Taxper,@oRecord),;
			_RecalcAll(@oRecord,@nProfit,nLoad),  ;
			dc_getrefresh(getlist)}

		@ 15,65 DCPUSHBUTTON CAPTION 'Get To;Payment 60' PARENT oTabpage3;
			SIZE 12,2;
			CARGO 'CANCEL' ;
			ACTION {|| upgettopay(60,oRecord:IntRate,@oRecord:FINANCE,oRecord:Pay60,@oRecord:DownPay,@oRecord:Quoted,oRecord:Taxper,@oRecord),;
			_RecalcAll(@oRecord,@nProfit,nLoad),   ;
			dc_getrefresh(getlist)}

	  @ 17,65 DCPUSHBUTTON CAPTION 'Get To;Payment 72' PARENT oTabpage3;
			SIZE 12,2;
			CARGO 'CANCEL' ;
			ACTION {|| upgettopay(72,oRecord:IntRate,@oRecord:FINANCE,oRecord:Pay72,@oRecord:DownPay,@oRecord:Quoted,oRecord:Taxper,@oRecord),;
			_RecalcAll(@oRecord,@nProfit,nLoad),   ;
			dc_getrefresh(getlist)}



		@ 19,65 DCPUSHBUTTON CAPTION 'Get To;Payment ODD' PARENT oTabpage3;
			SIZE 12,2;
			CARGO 'CANCEL' ;
			ACTION {|| upgettopay(oRecord:Termodd,oRecord:IntRate,@oRecord:FINANCE,oRecord:Payodd,@oRecord:DownPay,@oRecord:FINANCE,oRecord:Taxper,@oRecord),;
			_RecalcAll(@oRecord,@nProfit,nLoad), ;
			dc_getrefresh(getlist)}


	@ 7,80 DCPUSHBUTTON CAPTION 'Clear All Interest Rates' PARENT oTabpage3;
			SIZE 20,2;
			CARGO 'CANCEL' ;
			ACTION {|| oRecord:Rate72:=oRecord:Rate36:=oRecord:Rate48:=oRecord:Rate60:=0,DC_GETREFRESH(GETLIST)}


 @ 32,1 DCTOOLBAR oToolbar3 SIZE 60,1.5 BUTTONSIZE 15,1.5  PARENT oTabpage3

	 DCADDBUTTONXP CAPTION 'Print Quote '                              ;
		PARENT oToolbar3                                              ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
      PRTQUOTE(oRecord:QUOTEVIN )} ;
		CONFIG oConfig1

    DCADDBUTTONXP CAPTION 'Buyers Order '                              ;
		PARENT oToolbar3                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		BuyersOrder(NEWUPS->(RECNO()),.f.,@oRecord)}  ;
		CONFIG oConfig1


	DCADDBUTTONXP CAPTION 'Exit; No Save '                              ;
		PARENT oToolbar3                                                                            ;
		ACTION {|| DC_READGUIEVENT(DCGUI_EXIT_ABORT,GETLIST)} ;
		COLOR 1,BD_SALMON;
		CONFIG oConfig1

	DCADDBUTTONXP CAPTION 'Save and Exit '                              ;
		PARENT oToolbar3                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
      DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}  ;
		COLOR 1,BD_KEYLIME ;
		CONFIG oConfig1



@ 0,0 DCTABPAGE oTabpage4 CAPTION 'Lease Payments ';
		RELATIVE oTabpage3 ;
		size 150,34  ;
		GOTFOCUS{|| _RecalcAll(oRecord,@nProfit,nLoad),;
		DC_GETREFRESH(GETLIST)}


	/// get lease vars if any lease co has been chosen
	SELECT ACELEASE
	ACELEASE->(DBSEEK(oRecord:lcompany))
	IF FOUND()
	  ACELEASE->(dbgoto(RECNO()))
	  uploadlease(@oRecord,Acelease->(recno()))
	ELSE
		oRecord:addtopymnt:=0
	ENDIF


	@ 1,5 DCSAY 'Zipcode' Get  oRecord:ZIPCODE SAYCOLOR 9,0                PARENT oTabPage4 ;
		VALID{|| SEEKZIP(@oRecord),_GETSALESTAXPER(@oRecord),;
		IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),;
		DC_GETREFRESH(GETLIST),.T.};
		POPUP{|| BROWGETZIP(@oRecord),DC_GETREFRESH(GETLIST)};
		POPCAPTION 'ZipSearch ' POPWIDTH 100 POPFONT '10.Courier New' POPSTYLE 1

	@ 1,35 DCSAY 'City' Get oRecord:CITYSTATE  Saycolor 9,0   picture '@!' PARENT oTabPage4
	@ 1,70 DCSAY 'State' get oRecord:STATE                                 PARENT oTabPage4  GETCOLOR{|| IIF( empty(oRecord:state),{1,BD_MANGOCREME} ,{1,BD_KEYLIME} )  }
	@ 1,85 DCSAY 'County ' get oRecord:COUNTY                              PARENT oTabPage4
	@ 1,120 DCSAY 'Sales Tax % 'get oRecord:taxper PICTURE '99.9999'       PARENT oTabPage4  GETCOLOR{|| IIF( oRecord:taxper < .01,{1,BD_MANGOCREME} ,{1,BD_KEYLIME} )  }

	/// put sales tax here

	@ 2,5 DCSAY 'Lease For '+oRecord:lastname   PARENT oTabpage4
	@ 2,30 DCSAY 'Bank' get oRecord:lcompany  ;
	POPUP{|| uppicklease(@oRecord),DC_getrefresh(getlist)} ;
	POPCAPTION 'Search ' POPWIDTH 80 POPFONT '8.Courier New' POPSTYLE 1    ;
	PARENT oTabpage4

	@ 2,70 DCCHECKBOX lCaptax  Prompt 'Capitalize Taxes' ;
		ACTION{||  IIF( lCaptax,oRecord:captax:='Y' ,oRecord:captax:='N' ),;
		IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST)}   ;
      PARENT oTabpage4

  @ 2,90 DCCHECKBOX lCapfees  Prompt 'Capitalize Fees' ;
		ACTION{||  IIF( lCapfees,oRecord:capfees:='Y' ,oRecord:capfees:='N' ),;
		IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST)}   ;
      PARENT oTabpage4

  @ 2,110 DCCHECKBOX lProptax  Prompt 'Include Prop Tax' ;
		ACTION {||  IIF( lProptax,oRecord:extrachar3:='Y' ,oRecord:extrachar3:='N' ),;
		IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST)}   ;
      PARENT oTabpage4

  @ 2,130 DCCHECKBOX oRecord:aprlease  Prompt 'APR Lease' ;
		ACTION {|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST)}   ;
      PARENT oTabpage4


	@ 3.5,70 DCCHECKBOX lRecalc PROMPT 'Check Here To Recalc Lease After Every Entry' font '14.Courier New' color {|| IIF( lRecalc,{1,BD_KEYLIME} ,{1,BD_SALMON} )};
		   PARENT oTabpage4

	@ 3,5 DCSAY 'Lease Interest Rate' Get oRecord:LEASERATE1  valid{|| makemoneyfact(oRecord), IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),;
			DC_GETREFRESH(GETLIST,,,oTabpage4),.t.} Picture '999.999';
		  PARENT oTabpage4 Saycolor 9,0 Getcolor 10,0 Getfont '12.Courier'

	@ 3,40 DCSAY 'Lease BuyRate' get oRecord:lbuyrate picture '999.999' PARENT oTabpage4 ;
			valid{||  IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),;
			DC_GETREFRESH(GETLIST,,,oTabpage4),.t.} ;
			Saycolor 9,0 Getcolor 10,0 Getfont '12.Courier'

	@ 4,5 DCSAY ' Enter Money Factor' get oRecord:MoneyFact1   PICTURE '9.999999' PARENT oTabpage4   ;
   VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.};
	GETCOLOR{|| IIF( oRecord:Moneyfact1=0,{1,BD_MANGOCREME} ,{1,BD_KEYLIME} ) }

	@ 4,40 DCSAY '     Reserves' get oRecord:lreserves   PICTURE '99999.99' PARENT oTabpage4 EDITPROTECT{|| TRUE } NOTABSTOP

	@ 5,5 DCSAY '   Enter Lease Term' Get oRecord:LTERM Picture '999' Saycolor 9,0 Getcolor 10,0 Getfont '12.Courier' PARENT oTabpage4  ;
   VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.};
	GETCOLOR{|| IIF( oRecord:Lterm=0,{1,BD_MANGOCREME} ,{1,BD_KEYLIME} )  }


	@ 6,5 DCSAY '     Miles Per Year' get oRecord:LMILES   PICTURE '999999' PARENT oTabpage4 ;
		GETCOLOR{|| IIF( oRecord:Lmiles=0,{1,BD_MANGOCREME} ,{1,BD_KEYLIME} )  }

	@ 6,58 DCSAY ' Excess Miles Per Year' get oRecord:xcessmile   PICTURE '99999' PARENT oTabpage4 ;
      VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.}

	@ 6,95 DCSAY ' Cost Per Mile' get oRecord:costxcess   PICTURE '9.99' PARENT oTabpage4 ;
		when{|| IIF( oRecord:xcessmile > 0, TRUE ,.f.)};
		GETCOLOR{|| IIF( oRecord:costxcess=0,{1,BD_MANGOCREME} ,{1,BD_KEYLIME} )  }  ;
      VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.}

	@ 7,5 DCSAY '  Out The Door Cash' get oRecord:downpay  VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.};
		PARENT oTabpage4  PICTURE '99999.99' ;
		GETCOLOR{|| IIF( oRecord:downpay=0,{1,BD_BANANA} ,{1,BD_KEYLIME} )  }

	@ 8,5 DCSAY '            Rebates' get oRecord:LREBATE   PARENT oTabpage4 VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.};
		 PICTURE '99999'   ;
		 GETCOLOR{|| IIF( oRecord:LREBATE=0,{1,BD_BANANA} ,{1,BD_KEYLIME} )  }

	@ 9,5 DCSAY '        Trade Value' get oRecord:TRADEQUOTE VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),;
			CHKPRIC(oRecord,@nProfit,.f.),;
			DC_GETREFRESH(GETLIST),.t.};
	     	PICTURE '999999.99' PARENT oTabpage4

	 @ 9,40 DCSAY '   Trade ACV' Get oRecord:TRADEACV PARENT oTabPage4 Picture '999999.99' when{|| IIF( oRecord:TRADEQUOTE > 0,.t. ,.f. )};
			  VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),;
			  CHKPRIC(oRecord,@nProfit,.f.),;
			  DC_GETREFRESH(GETLIST),.t.};
			  GETCOLOR{|| IIF( oRecord:TRADEACV=0,{1,BD_MANGOCREME} ,{1,BD_KEYLIME} )  }


	@ 9,70 DCSAY '  Trade Lien' get oRecord:LIEN  PICTURE '999999.99'  ;
			 VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),;
			 CHKPRIC(oRecord,@nProfit,.f.),;
			 DC_GETREFRESH(GETLIST),.t.} ;
		    PARENT oTabpage4

	@ 10,5 DCSAY '               MSRP' get oRecord:QUOTEMSRP  VALID{|| MAKERESIDUAL(oRecord),IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),;
		DC_GETREFRESH(GETLIST),.t.};
		GETCOLOR{|| IIF( oRecord:QUOTEMSRP=0,{1,BD_MANGOCREME} ,{1,BD_KEYLIME} )  } ;
		 PARENT oTabpage4

	@ 11,5 DCSAY '       Adds to MSRP' get oRecord:MSRPADD  picture '99999.99'  VALID{|| MAKERESIDUAL(oRecord), IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.};
		 PARENT oTabpage4

	@ 12,5 DCSAY '   Enter Residual %' get oRecord:RESPER PICTURE '999' VALID{|| MAKERESIDUAL(oRecord), IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.} PARENT oTabpage4
	@ 12,40 DCSAY 'Residual Amt'   get oRecord:RESIDUAL PARENT oTabpage4  PICTURE '999999.99'  ;
	GETCOLOR{|| IIF( oRecord:RESPER=0,{1,BD_MANGOCREME} ,{1,BD_KEYLIME} )  }

	@ 13,5 DCSAY '      Selling Price' Get oRecord:QUOTED Picture "999999.99"   Saycolor 9,0 Getcolor 10,0 Getfont '12.Courier';
		 VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),;
		 CHKPRIC(oRecord,@nProfit,.f.),;
       DC_GETREFRESH(GETLIST),.t.}    PARENT oTabpage4;
		 GETCOLOR{|| IIF( oRecord:Lterm=0,{1,BD_MANGOCREME} ,{1,BD_KEYLIME} )  }

	@ 13,40 DCSAY '  Veh Profit' get nProfit PICTURE '999999.99' PARENT oTabpage4 EDITPROTECT{|| TRUE } NOTABSTOP


	@ 14,5 DCSAY '   Acquisition Fees' GET oRecord:ACQUISFEE picture '99999.99' PARENT oTabpage4 ;
		 VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.}

	@ 15,5 DCSAY '     Inspection Fee' GET oRecord:INSPECTION Picture '99999.99'  PARENT oTabpage4 ;
		 VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.}


	@ 16,5 DCSAY '       Registration' GET oRecord:REGFEE Picture '99999.99'         PARENT oTabpage4 ;
		VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.}


	@ 17,5 DCSAY ' Processing Doc Fee' GET oRecord:PROCFEE Picture '99999.99'      PARENT oTabpage4 ;
		VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.}


	@ 18,5 DCSAY '     Other Tax Fee 'GET oRecord:WASTEFEE Picture '99999.99'       PARENT oTabpage4 ;
		VALID{|| IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),DC_GETREFRESH(GETLIST),.t.}



	@ 18,49 DCGROUP oLRebatex CAPTION  'Lease Rebate Section' SIZE 100,9.5 PARENT oTabpage4

			@ 1,1 DCSAY 'FactCode' get oRecord:LREBCODE1 PARENT oLRebatex PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:LREBCODE1,@oRecord:LREBATE1,@oRecord:LREBDESC1),;
				DC_GETREFRESH(GETLIST),.T.}
			@ 2,1 DCSAY '  Amount' get oRecord:LREBATE1 picture '9999.99' PARENT oLRebatex ;//when{|| IIF( !empty(oRecord:LREBDESC1),.t. ,.f. )};
				valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),;
				IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),dc_getrefresh(getlist),.t.}
			@ 3,1 DCSAY ' Rebate1' get oRecord:LREBDESC1 GETID 'CREBDESC1' PARENT oLRebatex  LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}


			@ 1,25 DCSAY 'FactCode' get oRecord:LREBCODE2 PARENT oLRebatex PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:LREBCODE2,@oRecord:LREBATE2,@oRecord:LREBDESC2),;
				DC_GETREFRESH(GETLIST),.T.}
			@ 2,25 DCSAY '  Amount'  get oRecord:LREBATE2 PARENT oLRebatex picture '9999.99'; // when{|| IIF( !empty(oRecord:REBDESC2),.t. ,.f. )};
				valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),;
				IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),dc_getrefresh(getlist),.t.}
			@ 3,25 DCSAY ' Rebate2' get oRecord:LREBDESC2 PARENT oLRebatex LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}

			@ 1,50 DCSAY 'FactCode' get oRecord:LREBCODE3 PARENT oLRebatex PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:LREBCODE3,@oRecord:LREBATE3,@oRecord:LREBDESC3),;
				DC_GETREFRESH(GETLIST),.T.}
			@ 2,50 DCSAY '  Amount'  get oRecord:LREBATE3 PARENT oLRebatex picture '9999.99'; // when{|| IIF( !empty(oRecord:LREBDESC3),.t. ,.f. )};
				valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),dc_getrefresh(getlist),.t.}
			@ 3,50 DCSAY ' Rebate3' get oRecord:LREBDESC3   PARENT oLRebatex  LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}


			@ 1,75 DCSAY 'FactCode' get oRecord:LREBCODE4 PARENT oLRebatex PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:LREBCODE4,@oRecord:LREBATE4,@oRecord:LREBDESC4),;
				DC_GETREFRESH(GETLIST),.T.}
			@ 2,75 DCSAY '  Amount' get oRecord:LREBATE4 PARENT oLRebatex picture '9999.99'; //when{|| IIF( !empty(oRecord:LREBDESC4),.t. ,.f. )};
				valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),;
				IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),dc_getrefresh(getlist),.t.}
			@ 3,75 DCSAY ' Rebate4' get oRecord:LREBDESC4  PARENT oLRebatex   LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}


			@ 5,1 DCSAY 'FactCode' get oRecord:LREBCODE5 PARENT oLRebatex PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:LREBCODE5,@oRecord:LREBATE5,@oRecord:LREBDESC5),;
				DC_GETREFRESH(GETLIST),.T.}
			@ 6,1 DCSAY '  Amount' get  oRecord:LREBATE5 PARENT oLRebatex picture '9999.99'; //when{|| IIF( !empty(oRecord:LREBDESC5),.t. ,.f. )};
				valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),;
				IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),dc_getrefresh(getlist),.t.}
			@ 7,1 DCSAY ' Rebate5' get oRecord:LREBDESC5 PARENT oLRebatex   LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}

			@ 5,25 DCSAY 'FactCode' get oRecord:LREBCODE6 PARENT oLRebatex PICTURE '@!' VALID {|| _GETREBATEINFO(@oRecord:LREBCODE6,@oRecord:LREBATE6,@oRecord:LREBDESC6),;
				DC_GETREFRESH(GETLIST),.T.}
			@ 6,25 DCSAY '  Amount' get  oRecord:LREBATE6 PARENT oLRebatex picture '9999.99'; //when{|| IIF( !empty(oRecord:LREBDESC5),.t. ,.f. )};
				valid{|| _REFRESHREBATES(@oRecord),Addup(@oRecord),calccod(@oRecord),;
				IIF( lRecalc,_get2desiredups(oRecord,oTabpage4,.F.),nil),dc_getrefresh(getlist),.t.}
			@ 7,25 DCSAY ' Rebate6' get oRecord:LREBDESC6 PARENT oLRebatex   LOSTFOCUS{|| DC_GETREFRESH(GETLIST)}

			@ 8,1 DCSAY 'Total Lease Rebates' GET oRecord:LREBATE PARENT oLRebatex PICTURE '99999.99'




	@ 20,5 DCSAY '     Gross Cap Cost' get oRecord:extranum  picture '999999.99' editprotect{|| TRUE } PARENT oTabpage4
	@ 21,5 DCSAY '  Cash Down Applied' get oRecord:extranum2  picture '999999.99' editprotect{|| TRUE } PARENT oTabpage4
	@ 22,5 DCSAY 'Cap Cost Reductions' get oRecord:CAPCOSTRED  picture '999999.99' editprotect{|| TRUE } PARENT oTabpage4
	@ 23,5 DCSAY '  Adjusted Cap Cost' get oRecord:capcost  picture '999999.99' editprotect{|| TRUE } PARENT oTabpage4
	IF upper(oRecord:state)='CT'
	    @ 24,5 DCSAY '    Pre Tax Payment' get oRecord:lPayment picture '999999.99' editprotect{|| TRUE } PARENT oTabpage4
		 @ 25,5 DCSAY ' Sales Tax on Lease' get oRecord:LEASETAX  picture '999999.99' editprotect{|| TRUE } PARENT oTabpage4
	ENDIF
	@ 26,5 DCSAY '   Lease Inceptions' get oRecord:INCEPTS   picture '999999.99' editprotect{|| TRUE } PARENT oTabpage4
	//@ 26,5 DCSAY '     Residual Value' get oRecord:RESIDUAL  picture '999999.99' editprotect{|| TRUE } PARENT oTabpage4
	@ 27.5,5 DCSAY '    Monthly Payment' get oRecord:LPAYMENT2 PICTURE '999999.99'  PARENT oTabpage4      GETCOLOR 1,BD_LINEN;
		 GETFONT '12.Arial Bold'


	@ 29,5 DCSAY ' Payment Includes Property Tax Of' get oRecord:extranum3 PICTURE '999999.99' PARENT oTabpage4  GETCOLOR{|| IIF( alltrim(oRecord:extrachar3)='Y',{1,BD_SALMON} ,{1,BD_LINEN} )};
		WHEN {|| IIF(alltrim(oRecord:extrachar3)='Y',.t.,.f.)}


	@ 31,5 DCSAY 'First Pay+Downpay Due At Signing ' GET oRecord:Dueatsign  PICTURE '999999.99' PARENT oTabpage4

	@ 10,100 DCPUSHBUTTONXP CAPTION 'Calculate Final Payment;Based On Above Input ' ;
	SIZE 20,4;
	CARGO 'CANCEL' ;
	CONFIG oConfig1 ;
	PARENT oTabpage4 ;
	ACTION {|| _get2desiredups(oRecord,oTabpage4,.F.),;
	lRecalc:=.t.,;
	DC_GETREFRESH(GETLIST),.t.}


	@ 15,100 DCPUSHBUTTONXP CAPTION 'Get To Payment!' ;
	SIZE 20,2;
	CARGO 'CANCEL' ;
	CONFIG oConfig1 ;
	PARENT oTabpage4 ;
	ACTION {|| _gettoleasepayups(@oRecord,oTabpage4),dc_getrefresh(getlist) }


   @ 32,1 DCTOOLBAR oToolbar4 SIZE 90,1.5 BUTTONSIZE 15,1.5  PARENT oTabpage4

	 DCADDBUTTONXP CAPTION 'Print Quote '                              ;
		PARENT oToolbar4                                              ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
      PRTQUOTE(oRecord:QUOTEVIN ),oRecord:WRITEUP:=.T.};
		CONFIG oConfig1

    DCADDBUTTONXP CAPTION 'Buyers Order '                              ;
		PARENT oToolbar4                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		BuyersOrder(NEWUPS->(RECNO()),.f.,@oRecord)};
		CONFIG oConfig1

	 IF manager()
	 DCADDBUTTONXP CAPTION 'Final; Buyers Order '                              ;
		PARENT oToolCalc                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		BuyersOrder(NEWUPS->(RECNO()),.t.,@oRecord)}  ;
		CONFIG oConfig2

	 ENDIF


	DCADDBUTTONXP CAPTION 'Exit ;No Save '                              ;
		PARENT oToolbar4                                                                            ;
		ACTION {|| DC_READGUIEVENT(DCGUI_EXIT_ABORT,GETLIST)} ;
		COLOR 1,BD_SALMON ;
		CONFIG oConfig1



	DCADDBUTTONXP CAPTION 'Save and Exit '                              ;
		PARENT oToolbar4                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}  ;
		COLOR 1,BD_KEYLIME   ;
		CONFIG oConfig1



	DCADDBUTTONXP CAPTION 'InspObj' ;
        CARGO 'CANCEL' ;
        PARENT oToolbar4 ;
        ACTION {|| dc_arrayview(oRecord)  };
		  CONFIG oConfig1



IF MANAGER()

 @ 0,0 DCTABPAGE oTabpage5 CAPTION 'Aftersale/Warranty Info';
		RELATIVE oTabpage4 ;
		size 150,34
		//COLOR {|| {GRA_CLR_BLACK,BD_MANGOCREME}}


	SELECT FNIOPTS
	IF EMPTY(oRecord:AFTSALE1)
	 oRecord:AFTSALE1:=FNIOPTS->SOFT1
	ENDIF
	IF EMPTY(oRecord:AFTSALE2)
	 oRecord:AFTSALE2:=FNIOPTS->SOFT2
	ENDIF
	IF EMPTY(oRecord:AFTSALE3)
	 oRecord:AFTSALE3:=FNIOPTS->SOFT3
	ENDIF
	IF EMPTY(oRecord:AFTSALE4)
	 oRecord:AFTSALE4:=FNIOPTS->SOFT4
	ENDIF
	IF EMPTY(oRecord:AFTSALE5)
	 oRecord:AFTSALE5:=FNIOPTS->SOFT5
	ENDIF
	IF EMPTY(oRecord:AFTSALE6)
	 oRecord:AFTSALE6:=FNIOPTS->SOFT6
	ENDIF

	// SET UP AUTOMATICS
	IF oRecord:PROCFEE=0
	 oRecord:PROCFEE := FNIOPTS->Docfee
	ENDIF
	IF oRecord:NEWUSED='N'
	 IF oRecord:WASTEFEE=0
	  oRecord:WASTEFEE := FNIOPTS->Wastetire
	 ENDIF
	ENDIF
	//oRecord:INSPECTION:=FNIOPTS->Inspection
	//oRecord:WASTEFEE:=FNIOPTS->Wastetire

	/// figure cod etc
	//calccod(@oRecord)

	@ 2,1 DCSAY 'After Sale Items For Buyers Order' PARENT oTabpage5
	@ 4,1 DCSAY 'Service Contract' get oRecord:SERVCONT            PARENT oTabpage5
	@ 4,60 DCGET oRecord:SERVCONAM picture '@R 99999.99'  GETID('CONAMT') PARENT oTabpage5 ;
	 	VALID{|| CALCAFTSALETAX(@oRecord),calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}

	@ 5,1 DCSAY 'AfterSale Item1 ' get oRecord:AFTSALE1                   PARENT oTabpage5
	@ 5,60 DCGET oRecord:AFTSALEAM1 Picture '@R 99999.99'                 PARENT oTabpage5 ;
	VALID{|| CALCAFTSALETAX(@oRecord),calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}

	@ 6,1 DCSAY 'AfterSale Item2 ' get oRecord:AFTSALE2                   PARENT oTabpage5
	@ 6,60 DCGET oRecord:AFTSALEAM2 Picture '@R 99999.99'                 PARENT oTabpage5;
	  	VALID{|| CALCAFTSALETAX(@oRecord),calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}

	@ 7,1 DCSAY 'AfterSale Item3 ' get oRecord:AFTSALE3                   PARENT oTabpage5
	@ 7,60 DCGET oRecord:AFTSALEAM3 Picture '@R 99999.99'                 PARENT oTabpage5 ;
		VALID{|| CALCAFTSALETAX(@oRecord),calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}

	@ 8,1 DCSAY 'AfterSale Item4 ' get oRecord:AFTSALE4                   PARENT oTabpage5
	@ 8,60 DCGET oRecord:AFTSALEAM4 Picture '@R 99999.99'                 PARENT oTabpage5 ;
		VALID{|| CALCAFTSALETAX(@oRecord),calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}

	@ 9,1 DCSAY 'AfterSale Item5 ' get oRecord:AFTSALE5                   PARENT oTabpage5
	@ 9,60 DCGET oRecord:AFTSALEAM5 Picture '@R 99999.99'                 PARENT oTabpage5 ;
		VALID{|| CALCAFTSALETAX(@oRecord),calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}

	@ 10,1 DCSAY 'AfterSale6 Item ' get oRecord:AFTSALE6                  PARENT oTabpage5
	@ 10,60 DCGET oRecord:AFTSALEAM6 Picture '@R 99999.99'                PARENT oTabpage5 ;
		VALID{|| CALCAFTSALETAX(@oRecord),calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}

	@ 12,1 DCSAY 'Inspection Fee '                                        PARENT oTabpage5
	@ 12,50 DCGET oRecord:INSPECTION Picture '@R 99999.99'                PARENT oTabpage5 ;
		VALID{|| calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}

	@ 13,1 DCSAY 'Registration '                                          PARENT oTabpage5
	@ 13,50 DCGET oRecord:REGFEE Picture '@R 99999.99'                    PARENT oTabpage5 ;
		VALID{|| calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}

	@ 13,80 DCSAY 'Transfer Plates Y/N' get oRecord:REGTRANS PICTURE '@! L' PARENT oTabpage5
	@ 14,1 DCSAY 'Processing Fee '                                        PARENT oTabpage5
	@ 14,50 DCGET oRecord:PROCFEE Picture '@R 99999.99'                   PARENT oTabpage5 ;
		VALID{|| calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}

	@ 15,1 DCSAY 'Other Tax Fee '                                         PARENT oTabpage5
	@ 15,50 DCGET oRecord:WASTEFEE Picture '@R 99999.99'                  PARENT oTabpage5 ;
		VALID{|| calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}


	@ 16,1 DCSAY 'Verify County   ' get oRecord:COUNTY                    PARENT oTabpage5
	@ 16,60 DCSAY 'Verify Tax %    ' get oRecord:TAXPER picture '@R 99.999'  PARENT oTabpage5


	@ 17,1 DCSAY 'Verify Cash Down' get oRecord:DOWNPAY picture '@R 999999.99' PARENT oTabpage5;
		WHEN{|| IIF( !oRecord:QUOTETYPE='C',.T.,.F.)};
		VALID{|| calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}




	@ 18,1 DCSAY 'Verify Cash Deposit Submitted '  get oRecord:DEPOSIT PARENT oTabpage5 picture '@R 999999.99' ;
		VALID{|| calccod(@oRecord),DC_GETREFRESH(GETLIST),.T.}

	@ 18,80 DCSAY 'Customer Number' GET oRecord:REBDESC6 PICTURE '@!'  PARENT oTabpage5

	@ 19,1 DCSAY 'Cash Due On Delivery' get oRecord:COD picture '@R 999999.99' NOTABSTOP EDITPROTECT{|| TRUE } PARENT oTabpage5

	@ 20,1 DCSAY 'Enter Lien Bank Info/Acct#' get oRecord:LIENINFO PICTURE '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!'  PARENT oTabpage5
	@ 21,1 DCSAY 'Enter Insurance Agent Info' get oRecord:INSURANCE PICTURE '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!' PARENT oTabpage5
	@ 22,1 DCSAY 'Enter Registration Info   ' get oRecord:EXTRACHAR2 PICTURE '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!' PARENT oTabpage5
	@ 23,1 DCSAY 'Enter Delivery Date       ' get dDelivdate                                                            PARENT oTabpage5
	@ 23,60 DCSAY 'Delivery Time' get cDeltime   PARENT oTabpage5


	@ 24,0 DCMULTILINE oRecord:BOCOMMENT SIZE 60,4 FONT '8.Courier New' maxlines 4 LINELENGTH 60 ;
		  NOVERTSCROLL NOHORIZSCROLL PARENT oTabpage5

	@ 27,70 DCSAY 'InsuranceInfo' get oRecord:INSURANCE PICTURE '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!' getfont '8.Courier New'  PARENT oTabpage5


	@ 28.5,1 DCTOOLBAR oToolbar5 SIZE 100,1.5 BUTTONSIZE 15,1.5  PARENT oTabpage5
    DCADDBUTTONXP CAPTION 'Print Quote '                              ;
		PARENT oToolbar5                                              ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(@oRecord,dDelivdate,cDeltime),;
      PRTQUOTE(oRecord:QUOTEVIN ),oRecord:WRITEUP:=.T.};
		CONFIG oConfig1

    DCADDBUTTONXP CAPTION 'Buyers Order '                              ;
		PARENT oToolbar5                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(@oRecord,dDelivdate,cDeltime),;
		BuyersOrder(NEWUPS->(RECNO()),.t.,@oRecord)} ;
		CONFIG oConfig1

	  DCADDBUTTONXP CAPTION 'Final;Buyers Order '                              ;
		PARENT oToolbar5                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(oRecord,dDelivdate,cDeltime),;
		BuyersOrder(NEWUPS->(RECNO()),.t.,@oRecord)}   ;
		CONFIG oConfig1



	 DCADDBUTTONXP CAPTION 'Exit ;No Save '                              ;
		PARENT oToolbar5                                                                            ;
		ACTION {|| DC_READGUIEVENT(DCGUI_EXIT_ABORT,GETLIST)} ;
		COLOR 1,BD_SALMON  ;
		CONFIG oConfig1


	 DCADDBUTTONXP CAPTION 'Save and Exit '                              ;
		PARENT oToolbar5                                                                            ;
		ACTION {|| Savaquote(@cBody,@oRecord),;
		UPQTUPDT(@oRecord,dDelivdate,cDeltime),;
      DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}  ;
		COLOR 1,BD_KEYLIME ;
		CONFIG oConfig1



	@ 4,80 DCPUSHBUTTONXP CAPTION 'Search Extended Warranty' PARENT oTabpage5                          ;
		SIZE 28,1                                              ;
		action {|| browwarr(@oRecord:SERVCONT),;
		DC_GETREFRESH(GETLIST),SETAPPFOCUS(DC_GETOBJECT(GETLIST,'CONAMT'))};
		COLOR 1,BD_CLEARBLUE ;
		CONFIG oConfig1


	@ 6,80 DCSAY 'Get OptionSoft' GET cFileName PARENT oTabpage5;
  POPUP {|| cFilename:=DC_PopFile( ,'c:\optionsoft\deals\pending','*.xml'), ;
        GetOptionsoft(cFilename,@oRecord),ADDUP(@oRecord),DC_GETREFRESH(GETLIST)}

	@ 8,80 DCPUSHBUTTON CAPTION  'Make OptionSoft';
			SIZE 20,2 ;
		 	 PARENT oTabpage5;
		  ACTION {|| _OPTIONSOFT(NEWUPS->(RECNO()))}



 ENDIF
//wtf oRecord:bocomment, len(oRecord:bocomment) pause
DCREAD GUI TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA FIT OPTIONS GETOPTIONS TITLE 'QUOTE UP ENTRY'

//DCREAD GUI TO XSTATUS FIT ADDBUTTONS OPTIONS GETOPTIONS TITLE 'QUOTE UP ENTRY'
IF !XSTATUS
 Savaquote(@cBody,@oRecord)
 UPQTUPDT(@oRecord,dDelivdate,cDeltime)
 NEWUPS->(DBRUNLOCK(nUPRec))
 RETURN NIL
ENDIF
SELECT NEWUPS

//Savaquote(@cBody,@oRecord)
//UPQTUPDT(@oRecord,dDelivdate,cDeltime)

UPTAXTAB->(DBCLOSEAREA())
FNIOPTS->(DBCLOSEAREA())
SERVCONT->(DBCLOSEAREA())
/*
IF lOfferOC
	OILCHOPT->(DBCLOSEAREA())
ENDIF  */

NEWUPS->(DBRUNLOCK(nRec))
RETURN NIL

STATIC Function _SetAutoCharges(oRecord,cType)
	// SET UP AUTOMATICS

	IF cType='N'
	 IF oRecord:PROCFEE=0
	  oRecord:PROCFEE := FNIOPTS->Docfee
	 ENDIF
	 //IF oRecord:NEWUSED='N'
	  IF oRecord:WASTEFEE=0
	   oRecord:WASTEFEE := FNIOPTS->Wastetire
	  ENDIF
	ENDIF
	IF cType='U'
	 IF oRecord:PROCFEE=0
	  oRecord:PROCFEE := FNIOPTS->Docfee
	 ENDIF
	 oRecord:WASTEFEE:=0
	ENDIF


	RETURN NIL

static function _makeleasereserve(oRecord)
  LOCAL  cReserveBlk:='' , bBlock , nResper:=100
  /// do not attempt calculation is leasebuy rate is 0
  IF oRecord:lbuyrate < .001
	RETURN oRecord
  ENDIF
  /// get reserve calc code block if there is one
  SELECT ACELEASE
  ACELEASE->(DBSEEK(oRecord:LCOMPANY))
  IF ACELEASE->(FOUND())
	cReserveBlk:=alltrim(ACELEASE->RESCALC)
	nResper:=ACELEASE->RESPER
  ENDIF



  IF EMPTY(cReserveBlk)
	/// standard reserve calculation
	_Standlreserve(oRecord,nResper)

  else    /// use codeblock
	 IF cReserveblk='STANDARD'
		  _Standlreserve(oRecord,nResper)
		  return oRecord
	 ENDIF
	 /// WHATEVER IS IN THERE SAVE THE FIELDS FIRST FOR CODE BLOCK TO WORK
	NEWUPS->(SAVEFLDS(oRecord,{"CAPCOST","LTERM"}))
	cReserveBlk:= "{||"+cReserveBlk+"}"
	bBlock:=&(cReserveBlk)
	oRecord:lReserves:=EVAL(bBlock)
  endif


  RETURN oRecord

/// standard lease reserve calc that can be called as a code block from acelease resvcalc
static function _Standlreserve(oRecord,nResper)
  LOCAL nPayment:=0,nPayment2:=0 , nEms:=0 ,nDifference,r:=oRecord:leaserate1/1200 ,n:=oRecord:lBuyrate/1200

  //// calc payment sas an apr lease with lease rate
  nPayment:=(oRecord:capcost-(oRecord:Residual-nEms)/(1+r)^oRecord:lTerm)/((1-(1/(1+r)^oRecord:lterm))/r) // calc pay
  //// calc payment as an apr lease with lease buyrate
  nPayment2:=(oRecord:capcost-(oRecord:Residual-nEms)/(1+n)^oRecord:lTerm)/((1-(1/(1+n)^oRecord:lterm))/n) // calc pay


  nDifference:=(nPayment-nPayment2)*oRecord:lTerm

  oRecord:lreserves:=nDifference*(nResper/100)

	RETURN oRecord


static function _stranalyser(cStr)
	local x:="" , i   ,getlist[0]
	for i=1 to len(cStr)
		x+=alltrim(str(asc(cstr[i])))+"/"
	next
	@ 0,0 dcmultiline cStr size 110

	dcread gui fit
	return nil

static function _reinitstr(cStr)
	local cChar,  i, lReInit := TRUE

	FOR i=1 to len(cStr)
		IF isalpha(cstr[i]) .OR. isdigit(cStr[i])
			lReInit := FALSE
			exit
		ENDIF
	NEXT

	IF lreinit
		cstr:=''
	ENDIF
	return  TRUE


static FUNCTION uppicklease(oRecord)
	LOCAL GetList := {}, aPres, MoBrowse, MoToolBar
	local getoptions,xstatus , nRec:=0

	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE SAYFONT('10.Courier New') getfont('10.Courier New')
	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, -1 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 15 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

	@ 1,1 DCTOOLBAR MoToolBar                               ;
		SIZE 90, 1.5
	ACELEASE->(DBGOTOP())
/* ----- Create browse ----- */
	@ 3,1 DCBROWSE MoBrowse ALIAS 'ACELEASE'                 ;
		SIZE 110,25                                     ;
		FONT '10.Courier New'  ;
		PRESENTATION aPres ;
		DATALINK{|| nRec:=ACELEASE->(RECNO()),uploadlease(@oRecord,nRec),DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}

	DCBROWSECOL FIELD ACELEASE->LEASENO                     ;
		WIDTH 4 HEADER "Lease_ID" PARENT MoBrowse
	DCBROWSECOL FIELD ACELEASE->LEASENAME                     ;
		WIDTH 20 HEADER "LeaseCo" PARENT MoBrowse
	DCBROWSECOL FIELD ACELEASE->ACQFEE                     ;
		WIDTH 10 HEADER "Acquistion Fee " PARENT MoBrowse
	DCBROWSECOL FIELD ACELEASE->stdmiles                     ;
		WIDTH 5 HEADER "Miles/Year " PARENT MoBrowse
	DCBROWSECOL FIELD ACELEASE->city                     ;
		WIDTH 15 HEADER "City " PARENT MoBrowse
	DCBROWSECOL FIELD ACELEASE->state                     ;
		WIDTH 5 HEADER "State " PARENT MoBrowse
	DCBROWSECOL FIELD ACELEASE->notes                     ;
		WIDTH 20 HEADER "Notes " PARENT MoBrowse
	DCREAD GUI ;
		TO XSTATUS;
		FIT TITLE 'LEASECO BROWSER' ;
		OPTIONS GETOPTIONS;
		ADDBUTTONS      ;
		APPWINDOW MAINWINDOW():DRAWINGAREA
	IF !XSTATUS
		RETURN nil
	ENDIF
	nRec:=ACELEASE->(RECNO())
	upLOADLEASE(@oRecord,nRec)
RETURN oRecord

STATIC FUNCTION upLOADLEASE(oRecord,nRec)

	ACELEASE->(DBGOTO(nRec))
	oRecord:lCompany  := ACELEASE->leaseno
	oRecord:acquisfee    := ACELEASE->acqfee
	//IF oRecord:lMiles < 1
	// oRecord:lmiles := ACELEASE->stdmiles
	//ENDIF
	oRecord:costxcess := ACELEASE->stdovrate
	oRecord:aprlease   := ACELEASE->aprlease
	oRecord:addtopymnt:=ACELEASE->addtopymnt
	return nil






static function _WarnCtDeal(oRecord)
 IF UPPER(oRecord:state)='CT'
  IF oRecord:QUOTED > 49000
	 NANNA()
	 DC_WINALERT('Warning:  This is a Connecticut transaction that may require a higher tax rate!!!!! ')
	 xBROWCOUNTY(oRecord)

  ENDIF
 ENDIF
 RETURN TRUE

static function _getresidualperc(oRecord)
	IF oRecord:QUOTEMSRP+oRecord:MSRPADD > 0
     oRecord:RESPER:=(oRecord:Residual/(oRecord:QUOTEMSRP+oRecord:MSRPADD))*100
	endif
	RETURN oRecord

///this function recalculates ALL finance and lease payments
static function _RecalcAll(oRecord,nProfit,nLoad)
		IF valtype(oRecord:addtopymnt) <> 'N'
			oRecord:addtopymnt:=0
		ENDIF
	   Chkpric(@oRecord,@nProfit,.F.)
		Addup(@oRecord)
		calccod(@oRecord)
		calcpayment(60,oRecord:Rate60,@oRecord:Pay60,oRecord:FINANCE+nLoad)
	   calcpayment(48,oRecord:Rate48,@oRecord:Pay48,oRecord:FINANCE+nLoad)
	   calcpayment(36,oRecord:Rate36,@oRecord:Pay36,oRecord:FINANCE+nLoad)
	   calcpayment(72,oRecord:Rate72,@oRecord:Pay72,oRecord:FINANCE+nLoad)
	   calcpayment(oRecord:termodd,oRecord:IntRate,@oRecord:payodd,oRecord:FINANCE+nLoad)
		MAKERESIDUAL(@oRecord)
		_MAKECAPCOSTUps(@oRecord)
		RETURN oRecord


/// get sale tax percentage based on county
static function _getsalestaxper(oRecord)
	SELECT UPTAXTAB
	SEEK(oRecord:COUNTY)
	IF found()
		oRecord:taxper:=UPTAXTAB->TAXPER
	ENDIF
	return nil


/// refresh rebates
static function _REFRESHREBATES(oRecord)
	oRecord:REBATE:=oRecord:REBATE1+oRecord:REBATE2+oRecord:REBATE3+oRecord:REBATE4;
	+oRecord:REBATE5+oRecord:REBATE6
	oRecord:LREBATE:=oRecord:LREBATE1+oRecord:LREBATE2+oRecord:LREBATE3+oRecord:LREBATE4;
	+oRecord:LREBATE5+oRecord:LREBATE6
return oRecord

function _GETREBATEINFO(cRebcode,nRebate,cRebdesc)
	LOCAL getlist:={},getoptions,oRebRecord, xStatus
	DCGETOPTIONS SAYWIDTH 0 SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
	IF empty(cRebcode)
		RETURN TRUE
	ENDIF
	SELECT REBATEUP
	REBATEUP->(ORDSETFOCUS('REBATEUP'))
	REBATEUP->(DBSEEK(cRebcode))
	IF FOUND()
		//POPULATE VARS  THEN RETURN
		cRebdesc:=REBATEUP->REBDESC
		//nRebate:=REBATEUP->AMOUNT
		RETURN TRUE
	ENDIF
	/// ADD RECORD IF NOT FOUND
	oRebRecord:=REBATEUP->(DC_DBRECORD(): NEW())
	REBATEUP->(DB_INIT(oRebRecord))
	oRebRecord:REBCODE:=cRebcode
	@ 2,0 DCSAY 'Rebate Code to set up' get oRebRecord:rebcode  EDITPROTECT{|| TRUE }
	@ 4,0 DCSAY 'Description          ' get oRebRecord:rebdesc
	//@ 5,0 DCSAY 'Rebate Value         ' get oRebRecord:amount PICTURE '99999.99'
	dcread gui modal enterexit to xstatus options getoptions fit title 'Establish Rebate Record' eval{|o| setappwindow(o)}
	IF !xStatus
		return TRUE
	ENDIF


	REBATEUP->(DC_DBGATHER(oRebRecord,.t.))

	/// POPULATE VARS
	cRebdesc:=REBATEUP->REBDESC
	//nRebate:=REBATEUP->AMOUNT
	RETURN TRUE





static function APPTOLS1(oRecord,lApptols1)
 IF lApptols1
 	 oRecord:LREBDESC1:=oRecord:REBDESC1
	 oRecord:LREBATE1:=oRecord:REBATE1
	 oRecord:LREBCODE1:=oRecord:REBCODE1
	else
	   oRecord:LREBDESC1:=space(10)
		oRecord:LREBATE1:=0
		oRecord:LREBCODE1:=space(10)
 endif
return oRecord

static function APPTOLS2(oRecord,lApptols2)
 IF lApptols2
 	 oRecord:LREBDESC2:=oRecord:REBDESC2
	 oRecord:LREBATE2:=oRecord:REBATE2
	 oRecord:LREBCODE2:=oRecord:REBCODE2
	else
	   oRecord:LREBDESC2:=space(10)
		oRecord:LREBATE2:=0
		oRecord:LREBCODE2:=space(10)
 endif
return oRecord
static function APPTOLS3(oRecord,lApptols3)
 IF lApptols3
 	 oRecord:LREBDESC3:=oRecord:REBDESC3
	 oRecord:LREBATE3:=oRecord:REBATE3
	 oRecord:LREBCODE3:=oRecord:REBCODE3
	else
	   oRecord:LREBDESC3:=space(10)
		oRecord:LREBATE3:=0
		oRecord:LREBCODE3:=space(10)
 endif
return oRecord

static function APPTOLS4(oRecord,lApptols4)
 IF lApptols4
 	 oRecord:LREBDESC4:=oRecord:REBDESC4
	 oRecord:LREBATE4:=oRecord:REBATE4
	 oRecord:LREBCODE4:=oRecord:REBCODE4
	else
	   oRecord:LREBDESC4:=space(10)
		oRecord:LREBATE4:=0
		oRecord:LREBCODE4:=space(10)
 endif
return oRecord
static function APPTOLS5(oRecord,lApptols5)
 IF lApptols5
 	 oRecord:LREBDESC5:=oRecord:REBDESC5
	 oRecord:LREBATE5:=oRecord:REBATE5
	 oRecord:LREBCODE5:=oRecord:REBCODE5
	else
	   oRecord:LREBDESC5:=space(10)
		oRecord:LREBATE5:=0
		oRecord:LREBCODE5:=space(10)
 endif
 return oRecord

 static function APPTOLS6(oRecord,lApptols6)
 IF lApptols6
 	 oRecord:LREBDESC6:=oRecord:REBDESC6
	 oRecord:LREBATE6:=oRecord:REBATE6
	 oRecord:LREBCODE6:=oRecord:REBCODE6
	else
	   oRecord:LREBDESC6:=space(10)
		oRecord:LREBATE6:=0
		oRecord:LREBCODE6:=space(10)
 endif

 return oRecord




/*
static function makecapcost(oRecord)
LOCAL nGrosscap:=0, lAprlease:=.f. ,nBasepay:=0,nTaxCCR:=0,nTaxonpay:=0,nNYTax:=0
 nGrosscap := oRecord:QUOTED+oRecord:ACQUISFEE

 oRecord:capcost := (nGrosscap-(oRecord:downpay+(oRecord:tradequote-oRecord:lien)+oRecord:lrebate))

 nNYtax:=1/(1-(oRecord:taxper/100))


/// make basic untaxed payment
 makeBaseleasepay(@oRecord,@nBasepay)


// take 1st payment out of downpay and refigure lease


 nTaxCCR:=((oRecord:DOWNPAY-nBasepay)+oRecord:LREBATE)*oRecord:taxper/100

 nGrosscap:=nGrosscap+nBasepay+nTaxCCR
 oRecord:capcost := (nGrosscap-((oRecord:downpay)+(oRecord:tradequote-oRecord:lien)+oRecord:lrebate))

  makeBaseleasepay(@oRecord,@nBasepay)

/// calc tax on base pay and add to nGrosscap
 ntaxonpay:=nBasepay*oRecord:Lterm
 nTaxonpay:=nTaxonpay*(oRecord:taxper/100)
 nTaxonpay:=nTaxonpay*nNYtax
 nGrosscap:=nGrosscap+nTaxonPay

 oRecord:capcost := (nGrosscap-(oRecord:downpay+(oRecord:tradequote-oRecord:lien)+oRecord:lrebate))

 /// make final lease pay
 makeLeasePay(@oRecord)



oRecord:LEASETAX:=nTaxonpay
oRecord:INCEPTS:=oRecord:LPAYMENT+nTaxCCR
oRecord:CAPCOSTRED:=(oRecord:DOWNPAY-oRecord:INCEPTS)
oRecord:DUEATSIGN:=oRecord:DOWNPAY+oRecord:Lpayment

return TRUE

static Function makeBaseleasepay(oRecord,nBasepay)
local nPay1,nPay2,nTaxonpay, nLocalTax:=0,nTaxCCR:=0
local nEms:=0      /// extra miles charge
local nNYtax:=1.00
nPay1:=nPay2:=ntaxonpay:=0
nEms:=(oRecord:LTerm/12)*(oRecord:xcessmile*oRecord:Costxcess)



nPay1:=((oRecord:Capcost-(oRecord:RESIDUAL-nEms)) / oRecord:LTerm )
nPay2:=((oRecord:Capcost+(oRecord:RESIDUAL-nEms)) * (oRecord:Moneyfact1))
/// warn if lease is too high to cause a db error
if nPay1+nPay2 > 99999.99
	dc_winalert('The lease payment calculated is out of range. Please check factor.')
	return nil
endif

nBasepay:=nPay1+nPay2
return TRUE   */

static function makemoneyfact(oRecord)
if oRecord:leaserate1 > 0
	  oRecord:Moneyfact1 := oRecord:leaserate1/2400
endif
return TRUE


static function makeresidual(oRecord)
IF oRecord:Resper > 0
 oRecord:Residual :=(oRecord:Quotemsrp+oRecord:Msrpadd)*(oRecord:Resper/100)
endif
return TRUE

/*
static Function makeleasepay(oRecord)
local nPay1:=0,nPay2:=0, nLocalTax:=0
local nEms:=0,nTax1:=0      /// extra miles charge
local nNYtax:=1.00
nPay1:=nPay2:=0
nEms:=(oRecord:LTerm/12)*(oRecord:xcessmile*oRecord:Costxcess)
nPay1:=((oRecord:Capcost-(oRecord:RESIDUAL-nEms)) / oRecord:LTerm )
nPay2:=((oRecord:Capcost+(oRecord:RESIDUAL-nEms)) * (oRecord:Moneyfact1))

oRecord:LPAYMENT:=nPay1+nPay2
RETURN TRUE */



static function paybump(oRecord)
	local nBump:=0
	local getoptions, xStatus
	local getlist:={}
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
	@ 10,0 DCSAY 'You may Pad the payment by the entered amount'
	@ 12,0 DCSAY 'Enter Amount to Pad Payment by' get nBump picture '9999.99'
	dcread gui modal enterexit to xstatus options getoptions fit title 'Bump Payment' eval{|o| setappwindow(o)}
	IF !xStatus
		return nil
	ENDIF
	oRecord:PAY72:=oRecord:PAY72+nBump
	oRecord:PAY36:=oRecord:PAY36+nBump
	oRecord:PAY48:=oRecord:PAY48+nBump
	oRecord:PAY60:=oRecord:PAY60+nBump
	oRecord:PAYODD:=oRecord:PAYODD+nBump
 RETURN NIL


static function Addwaste(oRecord)
	IF oRecord:NEWUSED='N'
		oRecord:WASTEFEE:=FNIOPTS->WASTETIRE
	ENDIF
	RETURN NIL

STATIC FUNCTION UPQTUPDT(oRecord,dDelivdate,cDeltime)
	/// DO NOT WRITE NEG VALUES IN LPAYMENT
	// oNewUps
	IF oRecord:LPAYMENT < 0
		oRecord:LPAYMENT:=0
	ENDIF

	IF oRecord:LPAYMENT > 9999.99
		oRecord:LPAYMENT:=0
	ENDIF

	IF len(oRecord:Notes)  > 100000
		//DC_WINALERT('This is a bad Up Record- NOTES Please Call Bob Volz with this information-  Thanks')
		oRecord:Notes=''
	ENDIF

	IF len(oRecord:BOCOMMENT)  > 100000
		DC_WINALERT('This is a bad Up Record- BOCOMMENT Please Call Bob Volz with this information-  Thanks')
		oRecord:BOCOMMENT=''
	ENDIF


	//NEWUPS->(DBGOTO(oRecord:Record_Number))
	//NEWUPS->(DC_DBGATHER(oRecord))
	//NEWUPS->(DbGather)

	oRecord:delivdate := dDelivdate
	oRecord:deltime   := cDelTime
	oRecord:notes     := Alltrim(@oRecord:notes)
	oRecord:bocomment :=alltrim(@oRecord:bocomment)
	// NEWUPS->(SaveFlds(oRecord, {'delivdate',"deltime"}))


	NEWUPS->(SaveFlds(oRecord))


	//SLEEP(30)
	//IF NEWUPS->(DC_RECLOCK())
	// REPLACE NEWUPS->Delivdate with dDelivdate
	// REPLACE NEWUPS->Deltime with cDeltime
	// REPLACE NEWUPS->SM2 with oRecord:SM2
	//
	///ENDIF

  	//NEWUPS->(DBRUNLOCK(oRecord:Record_Number))

///ENDIF
return nil

static function CLEARTRADE(oRecord)
	local getlist:={}, getoptions, clCleartrade:='Y'  , xStatus
	DCGEToptions saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize TABSTOP
	@ 10,0 DCSAY 'Confirm: Do You Wish To Clear Trade Info' get clCleartrade picture '@! L'
	dcread gui modal enterexit ADDBUTTONS to xStatus options getoptions fit title 'Clear Trade' eval{|o| setappwindow(o)}
	IF !xStatus
		return nil
	ENDIF
	IF clCleartrade='Y'
	  oRecord:TRADEYR:=space(2)
	  oRecord:TRADEMAKE:=space(10)
	  oRecord:TRADEDESC:=space(10)
	  oRecord:TRADECOLOR:=space(10)
	  oRecord:TRADEMILES:=oRecord:TRADEACV:=oRecord:TRADEQUOTE:= 0
	  oRecord:TRADEVIN:=space(17)
	ENDIF
  RETURN NIL


static function _initCharArray(aChars)
ASize(aChars,0)
ASize(aChars,37) // or whatever the size should be
AFill(aChars,'')
RETURN NIL


static function _initNumArray(aNums)
ASize(aNums,0)
ASize(aNums,57) // or whatever the size should be
AFill(aNums,0)
RETURN NIL

static function _initUpsCharArray(aUpChars)
ASize(aUpChars,0)
ASize(aUpChars,27) // or whatever the size should be
AFill(aUpChars,'')
RETURN NIL



static function _initUpsNumArray(aUpNums)
ASize(aUpNums,0)
ASize(aUpNums,2) // or whatever the size should be
AFill(aUpNums,0)
RETURN NIL

static function _initUpSDATEArray(aUpDates)
ASize(aUpDates,0)
ASize(aUpDates,6) // or whatever the size should be
AFill(aUpDates,ctod('  /  /  '))
RETURN NIL


Static function BROWOCOPTS(oRecord,getlist)
	 IF !EMPTY(oRecord:EXTRACHAR)
		DC_GETREFRESH(GETLIST)
		RETURN TRUE
	 ENDIF
	 xBROWOCOPTS(@oRecord)
	 DC_GETREFRESH(GETLIST)
	 RETURN TRUE

static function xBROWOCOPTS(oRecord)
	LOCAL GetList := {}, aPres, MoBrowse, MoToolBar , xstatus
	LOCAL GETOPTIONS
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE sayfont '10.Courier New' getfont '10.Courier New'
	SELECT OILCHOPT
	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_ROWHEIGHT, -1},                  /* Row Height */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 15 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

	@ 1,1 DCTOOLBAR MoToolBar                               ;
		SIZE 60, 1.5   ;

		DCADDBUTTON CAPTION 'Add New Plan'                           ;
		SIZE 20 PARENT MoToolbar;
		ACTION {||addocplan(),MoBrowse:refreshall(),mobrowse:forcestable()};
		PARENT MoToolBar                                     ;
		TOOLTIP 'Add New Oil Change to file'

		DCADDBUTTON CAPTION 'Edit Plan'                           ;
		SIZE 20 PARENT MoToolbar;
		ACTION {||editocplan(),MoBrowse:refreshall(),mobrowse:forcestable()};
		PARENT MoToolBar                                     ;
		TOOLTIP 'Edit New Oil Change Plan'


/* ----- Create browse ----- */
	@ 3,1 DCBROWSE MoBrowse ALIAS 'OILCHOPT'                 ;
		SIZE 110,20                                     ;
		FONT '10.COURIER NEW'  ;
		PRESENTATION aPres ;
		DATALINK{|| oRecord:EXTRACHAR:=OILCHOPT->PLAN,oRecord:EXTRANUM:=OILCHOPT->COST,;
		DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}

	DCBROWSECOL FIELD OILCHOPT->PLAN                     ;
		WIDTH 15 HEADER "PLAN" PARENT MoBrowse

	DCBROWSECOL FIELD OILCHOPT->COST                     ;
		WIDTH 10 HEADER "Plan Cost" PARENT MoBrowse  ;
		PICTURE '9999.99'
	DCREAD GUI MODAL;
		TO XSTATUS;
		EVAL {|o| setappwindow(o),  SETAPPFOCUS(MobROWSE:GETCOLUMN(1))};
		FIT TITLE 'OIL CHANGE BROWSER' ;
		OPTIONS GETOPTIONS;
		BUTTONS DCGUI_BUTTON_OK

	IF !XSTATUS
		RETURN nil
	ENDIF
	oRecord:EXTRACHAR:=OILCHOPT->PLAN
	oRecord:EXTRANUM:=OILCHOPT->COST
RETURN TRUE


Static function _BROWCOUNTY(oRecord,getlist)
	 IF !EMPTY(oRecord:COUNTY)
		RETURN TRUE
	 ENDIF
	 xBROWCOUNTY(@oRecord)
	 DC_GETREFRESH(GETLIST)
	 RETURN TRUE

static function xBROWCOUNTY(oRecord)
	LOCAL GetList := {}, aPres, MoBrowse, MoToolBar , xstatus
	LOCAL GETOPTIONS
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE sayfont '10.Courier New' getfont '10.Courier New'
	SELECT UPTAXTAB
	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, -1 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 15 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

	@ 1,1 DCTOOLBAR MoToolBar                               ;
		SIZE 80, 1.5   ;

		DCADDBUTTON CAPTION 'Add New County'                           ;
		SIZE 20 PARENT MoToolbar;
		ACTION {||addtaxtab(),MoBrowse:refreshall(),mobrowse:forcestable()};
		PARENT MoToolBar                                     ;
		TOOLTIP 'Add New County'

		DCADDBUTTON CAPTION 'Edit County'                           ;
		SIZE 20 PARENT MoToolbar;
		ACTION {||edittaxtab(),MoBrowse:refreshall(),mobrowse:forcestable()};
		PARENT MoToolBar                                     ;
		TOOLTIP 'Edit County'

/* ----- Create browse ----- */
	@ 3,1 DCBROWSE MoBrowse ALIAS 'UPTAXTAB'                 ;
		SIZE 60,10                                     ;
		FONT '10.COURIER NEW'  ;
		PRESENTATION aPres ;
		DATALINK{|| oRecord:COUNTY:=UPTAXTAB->TAXENT,oRecord:TAXPER:=UPTAXTAB->TAXPER,;
		DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}

		DCBROWSECOL FIELD UPTAXTAB->TAXPER                     ;
		HEADER "Percentage" PARENT MoBrowse  ;
		PICTURE '99.999'

		DCBROWSECOL FIELD UPTAXTAB->TAXENT                     ;
		WIDTH 20 HEADER "County" PARENT MoBrowse

		 DCBROWSECOL FIELD UPTAXTAB->TAXCODE                     ;
		WIDTH 10 HEADER "State" PARENT MoBrowse



	DCREAD GUI MODAL;
		TO XSTATUS;
		EVAL {|o| setappwindow(o) };
		FIT TITLE 'TAX TABLE BROWSER' ;
		OPTIONS GETOPTIONS;
		ADDBUTTONS


	IF !XSTATUS
		RETURN NIL
	ENDIF

	oRecord:COUNTY:=UPTAXTAB->TAXENT
	oRecord:TAXPER:=UPTAXTAB->TAXPER
RETURN NIL

function BROWCOUNTY(oRecord)
	LOCAL GetList := {}, aPres, oBrowse, oToolBar , xstatus
	LOCAL GETOPTIONS
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE sayfont '10.Courier New' getfont '10.Courier New'

	IF USEDB('UPTAXTAB')
		UPTAXTAB->(DBGOTOP())
	ELSE
		RETURN NIL
	ENDIF


	SELECT UPTAXTAB

	SET DELETED ON

	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, -1 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 15 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

	@ 1,1 DCTOOLBAR oToolBar                               ;
		SIZE 60, 1.5   ;

		DCADDBUTTON CAPTION 'Add New County'                           ;
		SIZE 20 PARENT oToolbar;
		ACTION {||addtaxtab(),oBrowse:refreshall(),obrowse:forcestable()};
		PARENT oToolBar                                     ;
		TOOLTIP 'Add New County'

		DCADDBUTTON CAPTION 'Edit County'                           ;
		SIZE 20 PARENT oToolbar;
		ACTION {||edittaxtab(),oBrowse:refreshall(),obrowse:forcestable()};
		PARENT oToolBar                                     ;
		TOOLTIP 'Edit County'

		DCADDBUTTON CAPTION 'Delete County'                           ;
		SIZE 20 PARENT oToolbar;
		ACTION {|| UPTAXTAB->(BV_BLANK(.T.,.T.)),oBrowse:refreshall(),obrowse:forcestable()};
		PARENT oToolBar                                     ;
		TOOLTIP 'Delete County'


/* ----- Create browse ----- */
	@ 3,1 DCBROWSE oBrowse ALIAS 'UPTAXTAB'                 ;
		SIZE 40,20                                     ;
		FONT '10.COURIER NEW'  ;
		PRESENTATION aPres ;
		EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITDOWN ;



		DCBROWSECOL FIELD UPTAXTAB->TAXENT                     ;
		WIDTH 15 HEADER "COUNTY" PARENT oBrowse ;
		PICTURE '@!'

		DCBROWSECOL FIELD UPTAXTAB->TAXCODE                     ;
		HEADER "State" PARENT oBrowse  ;
		PICTURE '@!'

		DCBROWSECOL FIELD UPTAXTAB->TAXPER                     ;
		HEADER "PERCENTAGE" PARENT oBrowse  ;
		PICTURE '99.999'

	DCREAD GUI MODAL;
		TO XSTATUS;
		EVAL {|o| setappwindow(o) };
		FIT TITLE 'TAX TABLE BROWSER' ;
		OPTIONS GETOPTIONS;
		ADDBUTTONS

	IF !XSTATUS
		RETURN NIL
	ENDIF

RETURN NIL




FUNCTION CHKPRIC(oRecord,nProfit,lWarn)
	DEFAULT lWarn:=.f.
	nProfit:=0
	nProfit:=(oRecord:QUOTED-oRecord:QUOTETISS)-(oRecord:TRADEQUOTE-oRecord:TRADEACV)

	IF lWarn
	 IF nProfit < 0
		dc_winalert('This deal is behind Net! No way! Are You Sure???')
		return .t.
	 ENDIF
	ENDIF
RETURN .T.


STATIC FUNCTION CALCAFTSALETAX(oRecord)
	local nTaxable:=0
	IF FNIOPTS->DOCFEETAX='Y'
		nTaxable:=oRecord:PROCFEE
	ENDIF
	IF FNIOPTS->SOFT1TAX = 'Y'
		nTaxable:=nTaxable+oRecord:AFTSALEAM1
	ENDIF
	IF FNIOPTS->SOFT2TAX = 'Y'
		nTaxable:=nTaxable+oRecord:AFTSALEAM2
	ENDIF
	IF FNIOPTS->SOFT3TAX = 'Y'
		nTaxable:=nTaxable+oRecord:AFTSALEAM3
	ENDIF
	IF FNIOPTS->SOFT4TAX = 'Y'
		nTaxable:=nTaxable+oRecord:AFTSALEAM4
	ENDIF
	IF FNIOPTS->SOFT5TAX = 'Y'
		nTaxable:=nTaxable+oRecord:AFTSALEAM5
	ENDIF
	IF FNIOPTS->SOFT6TAX = 'Y'
		nTaxable:=nTaxable+oRecord:AFTSALEAM6
	ENDIF
	IF FNIOPTS->Svcontax = 'Y'
	 nTaxable:=nTaxable+oRecord:SERVCONAM
	endif
   oRecord:AFTSALETAX:=nTaxable*(oRecord:TAXPER/100)
return nil

static function calccod(oRecord)

 //IF oRecord:FINANCED ='Y' .OR. oRecord:QUOTETYPE='R'
IF oRecord:QUOTETYPE $('RL')
	 oRecord:COD:=oRecord:DOWNPAY-oRecord:DEPOSIT
	 oRecord:FINANCE:=(oRecord:QUOTED+oRecord:TAX)-oRecord:TRADEQUOTE+;
	 (oRecord:SERVCONAM+oRecord:AFTSALEAM1+oRecord:AFTSALEAM2+oRecord:AFTSALEAM3;
	+oRecord:AFTSALEAM4+oRecord:AFTSALEAM5+oRecord:AFTSALEAM6;
	 +oRecord:AFTSALETAX)-(oRecord:DOWNPAY+oRecord:REBATE-oRecord:LIEN);
   +oRecord:PROCFEE+;
	+oRecord:REGFEE;
   +oRecord:INSPECTION;
   +oRecord:WASTEFEE
ELSE
	 oRecord:DOWNPAY:=0
	 oRecord:COD:=(oRecord:QUOTED+oRecord:TAX )-(oRecord:TRADEQUOTE+oRecord:DEPOSIT)+;
	(oRecord:SERVCONAM+oRecord:AFTSALEAM1+oRecord:AFTSALEAM2+oRecord:AFTSALEAM3;
	+oRecord:AFTSALEAM4+oRecord:AFTSALEAM5+oRecord:AFTSALEAM6;
	 +oRecord:AFTSALETAX)-(oRecord:DOWNPAY+oRecord:REBATE-oRecord:LIEN);
   +oRecord:PROCFEE+;
	+oRecord:REGFEE;
   +oRecord:INSPECTION;
   +oRecord:WASTEFEE
	 oRecord:FINANCE:=0
 endif
return oRecord:COD

FUNCTION ADDUP(oRecord,nTotNonVeh)
  nTotNonVeh:=oRecord:AFTSALEAM1+oRecord:AFTSALEAM2+oRecord:AFTSALEAM3+oRecord:AFTSALEAM4;
	+oRecord:AFTSALEAM5+oRecord:AFTSALEAM6+oRecord:INSPECTION+oRecord:REGFEE+oRecord:PROCFEE;
	+oRecord:WASTEFEE+oRecord:SERVCONAM+oRecord:AFTSALETAX

  oRecord:FINANCE:=(oRecord:QUOTED-oRecord:TRADEQUOTE)
  IF oRecord:FINANCE > 0
  	oRecord:TAX:=(oRecord:FINANCE)*(oRecord:TAXPER/100)
  ELSE
  	oRecord:TAX:=0
  ENDIF
  oRecord:FINANCE:=oRecord:FINANCE+oRecord:TAX+oRecord:LIEN
  oRecord:FINANCE:=oRecord:FINANCE+nTotNonVeh-(oRecord:REBATE+oRecord:DOWNPAY)

RETURN oRecord:FINANCE


FUNCTION CHKADDUP(oRecord,nTotNonVeh)
	if oRecord:QUOTETYPE='C'
		oRecord:PAY72:=oRecord:PAY36:=oRecord:PAY48:=oRecord:PAY60:=oRecord:PAYODD:=oRecord:TERMODD:=oRecord:INTRATE:=0
		return .t.
	endif
	IF oRecord:FINANCE=0
		RETURN .T.
	ENDIF
	oRecord:FINANCE:=(oRecord:QUOTED-oRecord:TRADEQUOTE)
   IF oRecord:FINANCE > 0
   	oRecord:TAX:=oRecord:FINANCE*(oRecord:TAXPER/100)
   ELSE
  	 oRecord:TAX:=0
   ENDIF
   oRecord:FINANCE:=oRecord:FINANCE+oRecord:TAX+oRecord:LIEN
   oRecord:FINANCE:=oRecord:FINANCE+nTotNonVeh-(oRecord:REBATE+oRecord:DOWNPAY)

	Quick(@oRecord:FINANCE,@oRecord:PAY72,@oRecord:PAY36,@oRecord:PAY48,@oRecord:PAY60,;
		  @oRecord:PAYODD,@oRecord:TERMODD,@oRecord:INTRATE,@oRecord:DOWNPAY,@oRecord:QUOTED,.f.,;
		  oRecord:TAXPER,@oRecord:RATE72,@oRecord:RATE36,@oRecord:RATE48,@oRecord:RATE60)



	IF (oRecord:QUOTED-oRecord:TRADEQUOTE) > 0
		oRecord:TAX:=(oRecord:QUOTED-oRecord:TRADEQUOTE)*(oRecord:TAXPER/100)   // recalc tax
	ENDIF
RETURN .T.

/* BROWwarr PROGRAM TO DISPLAY SERVICE CONTRACT SELECTIONS
static FUNCTION browwarr(cServCont)
	LOCAL GetList := {}, aPres, MoBrowse, MoToolBar , xstatus
	LOCAL GETOPTIONS
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE sayfont '10.Courier New' getfont '10.Courier New'
	SELECT SERVCONT
	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },         ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY },      ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },       ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },       ;
    { XBP_PP_COL_DA_ROWHEIGHT, 15 },                       ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY },      ;
    { XBP_PP_COL_DA_CELLHEIGHT, 15 }  }


	@ 1,1 DCTOOLBAR MoToolBar                               ;
		SIZE 110, 1.5   ;

		DCADDBUTTON CAPTION 'Add New Contract'                           ;
		SIZE 20 PARENT MoToolbar;
		ACTION {||addsvcont(),MoBrowse:refreshall(),mobrowse:forcestable()};
		PARENT MoToolBar                                     ;
		TOOLTIP 'Add New Contract to file'



	@ 3,1 DCBROWSE MoBrowse ALIAS 'SERVCONT'                 ;
		SIZE 110,20                                     ;
		FONT '10.COURIER NEW'  ;
		PRESENTATION aPres ;
		DATALINK{cServCont:=alltrim(Servcont->vendor)+' '+servcont->descript,;
		DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)};
		EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITEXIT

	DCBROWSECOL FIELD SERVCONT->VENDOR                     ;
		WIDTH 10 HEADER "VENDOR" PARENT MoBrowse
	DCBROWSECOL FIELD SERVCONT->SERVCONT                     ;
		WIDTH 20 HEADER "Plan #" PARENT MoBrowse
	DCBROWSECOL FIELD SERVCONT->DESCRIPT                     ;
		WIDTH 40 HEADER "DESCRIPTION" PARENT MoBrowse
	DCBROWSECOL FIELD SERVCONT->COST                     ;
		WIDTH 7 HEADER "COST"  PICTURE '9999.99' PARENT MoBrowse
	DCBROWSECOL FIELD SERVCONT->SCTERM                     ;
		WIDTH 3 HEADER "TERM/MOS"  PICTURE '999' PARENT MoBrowse
	DCBROWSECOL FIELD SERVCONT->SCMILES                     ;
		WIDTH 6 HEADER "TERM/MILES"  PICTURE '999999' PARENT MoBrowse
	DCBROWSECOL FIELD SERVCONT->DEDUCT                     ;
		WIDTH 6 HEADER "DEDUCTIBLE"  PICTURE '9999' PARENT MoBrowse

	DCREAD GUI MODAL;
		TO XSTATUS;
		EVAL {|o| setappwindow(o),  SETAPPFOCUS(MobROWSE:GETCOLUMN(1))};
		FIT TITLE 'CONTRACT BROWSER' ;
		OPTIONS GETOPTIONS;
		BUTTONS DCGUI_BUTTON_OK

	IF !XSTATUS
		RETURN nil
	ENDIF
	cServCont:=alltrim(Servcont->vendor)+' '+servcont->descript
RETURN nil         */

FUNCTION BROWSM()
	LOCAL GetList := {}, aPres, oBrowse, oToolBar , xstatus
	LOCAL GETOPTIONS
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE sayfont '10.Courier New' getfont '10.Courier New'
	SET DELETED ON

	IF !UseDb({'SM'})
		 DC_WINALERT('Cannot Open Files     .')
	   	 RETURN NIL
	ENDIF
	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, -1 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 15 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

	@ 1,1 DCTOOLBAR oToolBar                               ;
		SIZE 60, 1.5   ;

		DCADDBUTTON CAPTION 'Add Sales Person'                           ;
		SIZE 20 PARENT oToolbar;
		ACTION {||smadd(),DC_GETREFRESH(GETLIST),oBrowse:refreshall(),obrowse:forcestable()};
		PARENT oToolBar                                     ;
		TOOLTIP 'Add New Salesperson'

		DCADDBUTTON CAPTION 'Delete Sales Person'                           ;
		SIZE 20 PARENT oToolbar;
		ACTION {||smdelete(RECNO()),DC_GETREFRESH(GETLIST),oBrowse:refreshall(),obrowse:forcestable()};
		PARENT oToolBar                                     ;
		TOOLTIP 'Delete Salesperson'

		DCADDBUTTON CAPTION 'Print List'                           ;
		SIZE 20 PARENT oToolbar;
		ACTION {||smList(),oBrowse:refreshall(),obrowse:forcestable()};
		PARENT oToolBar                                     ;
		TOOLTIP 'Print Salesperson List'


/* ----- Create browse ----- */
	@ 5,1 DCBROWSE oBrowse ALIAS 'SM'                 ;
		SIZE 110,20                                     ;
		FONT '10.COURIER NEW'  ;
		PRESENTATION aPres ;
		EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITEXIT

	DCBROWSECOL FIELD SM->SALESMAN                     ;
		WIDTH 5 HEADER "SP Initials" PARENT oBrowse HCOLOR 1,5
	DCBROWSECOL FIELD SM->NAME                     ;
		WIDTH 20 HEADER "Name" PARENT oBrowse HCOLOR 1,5
	DCBROWSECOL FIELD SM->INETNAME                     ;
		WIDTH 20 HEADER "Inet Name" PARENT oBrowse HCOLOR 1,5


	DCBROWSECOL FIELD SM->ID                     ;
		WIDTH 5 HEADER "Employee#" PARENT oBrowse HCOLOR 1,5
	DCBROWSECOL FIELD SM->ACTIVE                     ;
		WIDTH 2 HEADER "Active" PARENT oBrowse  HCOLOR 1,5
	DCBROWSECOL FIELD SM->SSN                     ;
		WIDTH 5 HEADER "SSN ID" PARENT oBrowse  HCOLOR 1,5

	DCREAD GUI ;
		TO XSTATUS;
		EVAL {|o| setappwindow(o)};
		FIT TITLE 'Sales Person Browser' ;
		OPTIONS GETOPTIONS;
		BUTTONS DCGUI_BUTTON_OK

	IF !XSTATUS
		RETURN nil
	ENDIF
	CLOSE ALL
RETURN nil

FUNCTION BROWADSOURCE()
	LOCAL GetList := {}, aPres, oBrowse, oToolBar , xstatus
	LOCAL GETOPTIONS
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE sayfont '10.Courier New' getfont '10.Courier New'
	IF !UseDb({'ADSOURCE'})
		 DC_WINALERT('Cannot Open Files     .')
	   	 RETURN NIL
	ENDIF


	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, -1 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 15 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

	@ 1,1 DCTOOLBAR oToolBar                               ;
		SIZE 60, 1.5   ;

		DCADDBUTTON CAPTION 'Add Adv Source'                           ;
		SIZE 20 PARENT oToolbar;
		ACTION {||adsourceadd(),DC_GETREFRESH(GETLIST),oBrowse:refreshall(),obrowse:forcestable()};
		PARENT oToolBar                                     ;
		TOOLTIP 'Add New Advertising Source'

		DCADDBUTTON CAPTION 'Delete Adv Source'                           ;
		SIZE 20 PARENT oToolbar;
		ACTION {||adsourcedelete(RECNO()),DC_GETREFRESH(GETLIST),oBrowse:refreshall(),obrowse:forcestable()};
		PARENT oToolBar                                     ;
		TOOLTIP 'Delete Advertising Source'

		DCADDBUTTON CAPTION 'Print List'                           ;
		SIZE 20 PARENT oToolbar;
		ACTION {||listads(),oBrowse:refreshall(),obrowse:forcestable()};
		PARENT oToolBar                                     ;
		TOOLTIP 'Print Advertising Source'


/* ----- Create browse ----- */
	@ 3,1 DCBROWSE oBrowse ALIAS 'ADSOURCE'                 ;
		SIZE 110,20                                     ;
		FONT '10.COURIER NEW'  ;
		PRESENTATION aPres ;
		EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITEXIT

	DCBROWSECOL FIELD ADSOURCE->ADSRCE                     ;
		WIDTH 5 HEADER "Source" PARENT oBrowse HCOLOR 1,5
	DCBROWSECOL FIELD ADSOURCE->DESC                     ;
		WIDTH 20 HEADER "Description" PARENT oBrowse HCOLOR 1,5

	DCREAD GUI MODAL;
		TO XSTATUS;
		EVAL {|o| setappwindow(o),  SETAPPFOCUS(obROWSE:GETCOLUMN(1))};
		FIT TITLE 'Advertising Source Browser' ;
		OPTIONS GETOPTIONS;
		BUTTONS DCGUI_BUTTON_OK

	IF !XSTATUS
		RETURN nil
	ENDIF
	CLOSE ALL
RETURN nil






static function addsvcont()
	local getlist:={}
	local getoptions
	DCGEToptions saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	if servcont->(dc_addrec(3,.t.,2))
		@ 10,0 DCSAY 'Vendor           ' get servcont->vendor
		@ 11,0 DCSAY 'Plan Number      ' get servcont->servcont
		@ 12,0 DCSAY 'Plan Description ' get servcont->descript
		@ 13,0 DCSAY 'Term in Months   ' get servcont->scterm
		@ 14,0 DCSAY 'Term in Miles    ' get servcont->scmiles
		dcread gui modal enterexit options getoptions title 'Add Service Contract Type' fit ;
			eval{|o| setappwindow(o)}
		//servcont->(dbcommit())
	endif
return nil

static function addocplan()
	local getlist:={}
	local getoptions
	DCGEToptions saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	if OILCHOPT->(dc_addrec())
		@ 10,0 DCSAY 'Plan           ' get OILCHOPT->PLAN
		@ 11,0 DCSAY 'Plan Cost      ' get OILCHOPT->COST
		dcread gui modal enterexit options getoptions title 'Add OC PLAN' fit ;
			eval{|o| setappwindow(o)}
		//servcont->(dbcommit())
	endif
return nil

static function editocplan()
	local getlist:={}
	local getoptions
	DCGEToptions saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	if OILCHOPT->(dc_reclock())
		@ 10,0 DCSAY 'Plan           ' get OILCHOPT->PLAN
		@ 11,0 DCSAY 'Plan Cost      ' get OILCHOPT->COST
		dcread gui modal enterexit options getoptions title 'Add OC PLAN' fit ;
			eval{|o| setappwindow(o)}
	endif
return nil

static function addtaxtab()
	local getlist:={} ,cCounty:=space(15),cState:=space(2),nTaxper:=0, xStatus
	local getoptions
	DCGEToptions saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	@ 10,0 DCSAY 'County         ' get cCounty PICTURE '@!'
	@ 11,0 DCSAY 'Sales Tax Rate ' get nTaxper picture '99.999'
	@ 12,0 DCSAY 'State          ' get cState picture '@!'
	dcread gui modal to xStatus enterexit options getoptions title 'Add Tax ' fit ;
			eval{|o| setappwindow(o)}

	IF !xStatus
		return nil
	ENDIF
	IF UPTAXTAB->(DC_ADDREC(3,.T.,2))
		REPLACE UPTAXTAB->TAXENT WITH cCounty
		REPLACE UPTAXTAB->TAXPER WITH nTaxper
		REPLACE UPTAXTAB->TAXCODE WITH cState
		UPTAXTAB->(DBRUNLOCK())
	ENDIF
return nil

static function edittaxtab()
	local getlist:={}
	local getoptions
	DCGEToptions saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	if UPTaxTab->(dc_reclock())
		@ 9, 0 DCSAY '         State' get UPTAXTAB->TAXCODE PICTURE '@!'
		@ 10,0 DCSAY '        County' get UPTaxtab->taxent  PICTURE '@!'
		@ 11,0 DCSAY 'Sales Tax Rate' get UPTaxtab->Taxper  PICTURE '99.9999'
		dcread gui modal enterexit options getoptions title 'Edit Tax Table' fit ;
			eval{|o| setappwindow(o)}
	endif
	UPTAXTAB->(DBRUNLOCK())
return nil




FUNCTION VEHBROW(cNEW,cDemodstkNo)
	LOCAL aPush
	IF cNEW='N'
		aPush:=DB_PUSH('NCI',NCI->(ORDSETFOCUS()))
		SBROWBYDESC(@cDemodstkno)
		DB_POP(aPush)
		NCI->(DBGOTOP())
	ENDIF
	IF cNEW='U'
		aPush:=DB_PUSH('UCL01',UCL01->(ORDSETFOCUS()))
		SUBROWSTK(@cDemodstkno)
		select ucl01
		DB_POP(aPush)
	ENDIF
RETURN NIL

FUNCTION CHKQTTYP(cQuoteType)                   /// CHECK TO SEE IF QUOTETYPE IS R OR L
	IF cQuoteType $('RLC')
		RETURN .T.
	ENDIF
RETURN .F.

FUNCTION PRTQUOTE(cVehicleID)
	LOCAL GETLIST:={} , xStatus,oPrinter, nCopies:=1, lPDF:=.f.
	local getoptions , oBumpGroup ,lSideBySide:=.t. , lPrintRange:=.F. ,oParamGroup ,lShowPropTax:=.t.
	LOCAL cFont1,cFont2,cFont3,cFont4, lPrintPaywords:=.f.
	local nWalkbump:=0
	local cRange72Desc:='',cRange36Desc:='',cRange48Desc:='',cRange60Desc:='',cRangeOddDesc:=''
	LOCAL nRange72:=0,nRange36:=0,nRange48:=0,nRange60:=0,nRangeodd:=0
	DCGETOPTIONS SAYWIDTH 0 SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE TABSTOP
	/// establish box printing properties

	 @ 2,0 DCGROUP oBumpGroup CAPTION 'Payment Bump Parameters ' SIZE 100,9
	 @ 1,1 DCSAY 'You may enter a percentage amount to bump the payment to for payment range presentation.' PARENT oBumpGroup
	 @ 2,1 DCSAY 'The default range bump is 0%'                                                             PARENT oBumpGroup
	 @ 3,1 DCSAY 'This will print a payment range for each term on the quote.'                              PARENT oBumpGroup
	 @ 4,1 DCSAY 'Example: The calculated payment for 24 Months is 250 and you enter a bump amount of 5%. ' PARENT oBumpGroup
	 @ 5,1 DCSAY 'The displayed payment would be $262.50'                                                   PARENT oBumpGroup
	 @ 6,1 DCSAY 'If You Select Print Payment Range The System will Deduct/Add -$5 / +$20 for 24Mos.'       PARENT oBumpGroup
	 @ 7,1 DCSAY '-$10 / +$25 for 36Mos  -$15 / +$30 for 48Mos -$20 / +$35 for 60  -$20 / +$40 for 84Mos.'  PARENT oBumpGroup


	 @ 8,1 DCSAY 'The other Monthly terms will display in the same manner. '                                PARENT oBumpGroup
	 @ 9,1 DCSAY 'Enter bump percentage amount to add to payment for printing. E.G. 5 or 10' get nWalkbump picture '99'  PARENT oBumpGroup

	 @ 10,0 DCGROUP oParamGroup CAPTION 'Printing Parameters'  size 100,8
	 @ 1,1  DCCHECKBOX lSideBySide  PROMPT 'Print Both Retail and Lease Side By Side- Uncheck For No Side By Side '  PARENT oParamGroup
	 @ 2,1  DCCHECKBOX lPrintRange  PROMPT 'Print Payment Range Instead of Actual Dollar Values' PARENT oParamGroup
	 @ 3,1  DCCHECKBOX lPrintPayWords  PROMPT 'Express Payment Range in Words Rather than Dollar Values' PARENT oParamGroup
	 @ 4,1  DCCHECKBOX lShowPropTax  PROMPT 'Print Property Tax $ Amount on Lease If Applicable' PARENT oParamGroup
	 @ 5,1  DCSAY 'Number of Copies to Print ' get nCopies PICTURE '9'      PARENT oParamGroup

	 DCREAD GUI MODAL TO xStatus ADDBUTTONS FIT OPTIONS GETOPTIONS TITLE ' ' EVAL{|o| SETAPPWINDOW(o)}
	/// bump all payments by minimum
	IF nWalkbump > 0
	 nWalkbump:=(100+nWalkbump)/100
	 nRange72:=NEWUPS->PAY72*nWalkbump
	 nRange36:=NEWUPS->PAY36*nWalkbump
	 nRange48:=NEWUPS->PAY48*nWalkbump
	 nRange60:=NEWUPS->PAY60*nWalkbump
	 nRangeOdd:=NEWUPS->PAYodd*nWalkbump
	else
	 nRange72:=NEWUPS->PAY72
	 nRange36:=NEWUPS->PAY36
	 nRange48:=NEWUPS->PAY48
	 nRange60:=NEWUPS->PAY60
	 nRangeOdd:=NEWUPS->PAYodd
	ENDIF
	 /// establish the verbiage range   ie low Two Hundreds
	 IF lPrintRange
	  _getRangeDescX(@nRange36,@cRange36Desc,36)
	  _getRangeDescX(@nRange48,@cRange48Desc,48)
	  _getRangeDescX(@nRange60,@cRange60Desc,60)
	  _getRangeDescX(@nRange72,@cRange72Desc,72)
	  _getRangeDescX(@nRangeOdd,@cRangeOddDesc,NEWUPS->termodd)
	  /// switch to words
	  IF lPrintPayWords
	   _getRangeDesc(@nRange36,@cRange36Desc,36)
	   _getRangeDesc(@nRange48,@cRange48Desc,48)
	   _getRangeDesc(@nRange60,@cRange60Desc,60)
	   _getRangeDesc(@nRange72,@cRange72Desc,72)
	   _getRangeDesc(@nRangeOdd,@cRangeOddDesc,NEWUPS->termodd)
	  ENDIF
	 endif

	DCPRINT ON SIZE 66,132 TO oPrinter FONT '10.Courier New' PREVIEW
	IF VALTYPE(OPRINTER) # 'O' .OR. !OPRINTER:LACTIVE
				RETURN NIL
	ENDIF


	DO WHILE nCopies > 0
		if file('uplogo.bmp')
			@ 0,0,8,120 dcprint bitmap (GETFLG('ACCTPATH')+'uplogo.bmp')
		endif
		SELECT NEWUPS
		@ 9,0,17,130 DCPRINT BOX
		@ 10,5 DCPRINT SAY 'Customer Price Quotation' FONT '14.Courier New Bold'
		@ DC_PRINTERROW()+1,5 DCPRINT SAY NEWUPS->FIRSTNAME
		@ DC_PRINTERROW(),35  DCPRINT SAY NEWUPS->LASTNAME
		@ DC_PRINTERROW()+1,5 DCPRINT SAY NEWUPS->STREET
		@ DC_PRINTERROW()+1,5 DCPRINT SAY NEWUPS->CITYSTATE
		@ DC_PRINTERROW(),25 DCPRINT SAY NEWUPS->STATE
		@ DC_PRINTERROW(),32 DCPRINT SAY NEWUPS->ZIPCODE
		@ DC_PRINTERROW()+1,5 DCPRINT SAY 'Telephone Home/Work'
		@ DC_PRINTERROW(),DC_PRINTERCOL()+15 DCPRINT SAY NEWUPS->AREA
		@ DC_PRINTERROW(),DC_PRINTERCOL()+5 DCPRINT SAY NEWUPS->UPID
		@ DC_PRINTERROW(),DC_PRINTERCOL()+8 DCPRINT SAY '/ '+NEWUPS->WORKPHN

		@ 18,0,32,130 DCPRINT BOX

		@ 19.5,1 DCPRINT SAY 'Vehicle And Price Information:' FONT '14.COURIER BOLD'

		@ 21,1 DCPRINT SAY 'Stock Number '
		@ 21,20 DCPRINT SAY NEWUPS->STKQUOTED
		@ 21,30 DCPRINT SAY 'Last 6 Vin'
		@ 21,50 DCPRINT SAY SUBSTR(cVehicleID,12,6)

		@ 22,1 DCPRINT SAY NEWUPS->YRQUOTED
		@ 22,15 DCPRINT SAY NEWUPS->VEHQUOTED
		@ 22,55 DCPRINT SAY 'Mileage'
		@ 22,65 DCPRINT SAY  NEWUPS->QUOTEMILES PICTURE '999999'
		@ 23,1 DCPRINT SAY 'MSRP'
		@ 23,30 DCPRINT SAY NEWUPS->QUOTEMSRP PICTURE '$999,999.99'

		@ 24,1 DCPRINT SAY 'Your Price'
		@ 24,30 DCPRINT SAY NEWUPS->QUOTED PICTURE '$999,999.99'
		@ 25,1 DCPRINT SAY 'PreEstimated Fees'
		@ 25,30 DCPRINT SAY NEWUPS->regfee+NEWUPS->procfee+NEWUPS->wastefee+NEWUPS->inspection picture '$999,999.99'
		@ 26,1 DCPRINT SAY 'Trade Allowance'
		@ 26,30 DCPRINT SAY NEWUPS->TRADEQUOTE  PICTURE '$999,999.99'
		@ 26,50 DCPRINT SAY 'Trade Vehicle'
		@ 26,80 DCPRINT SAY  NEWUPS->TRADEYR+' '+NEWUPS->TRADEDESC
		@ 26,110 DCPRINT SAY 'Mileage'
		@ 26,120 DCPRINT SAY NEWUPS->TRADEMILES PICTURE '999999'
		@ 27,1 DCPRINT SAY 'Sales Tax '
		@ 27,30 DCPRINT SAY NEWUPS->Tax+NEWUPS->aftsaletax PICTURE '$999,999.99'
		@ 27,50  DCPRINT SAY 'Tax Rate ' +NEWUPS->COUNTY+' '+STR(NEWUPS->TaxPER)
		@ 28,1 DCPRINT SAY 'Lien Payoff'
		@ 28,30  DCPRINT SAY NEWUPS->LIEN  picture '$999,999.99'
		@ 29,1 DCPRINT SAY 'Down Payment'
		@ 29,30 DCPRINT SAY NEWUPS->DOWNPAY  picture '$999,999.99'
		@ 30,1  DCPRINT SAY 'Net Amount Due'
		@ 30,30 DCPRINT SAY (NEWUPS->QUOTED+NEWUPS->TAX+NEWUPS->AFTSALETAX+NEWUPS->LIEN)-(NEWUPS->REBATE+NEWUPS->DOWNPAY+NEWUPS->TRADEQUOTE)  picture '$999,999.99'


		@ 33,0,44.2,64 DCPRINT BOX
		@ 33,65,44.2,130 DCPRINT BOX


		@ 35,1 DCPRINT SAY 'Retail Rebates:   '
		@ 35,66 DCPRINT SAY 'Lease Rebates:   '


		IF NEWUPS->REBATE1 > 0
			@ 36,1 DCPRINT SAY NEWUPS->REBDESC1
			@ 36,20 DCPRINT SAY NEWUPS->REBCODE1
			@ 36,30 DCPRINT SAY NEWUPS->REBATE1 PICTURE '$99,999.99'
		ENDIF

		IF NEWUPS->LREBATE1 > 0
			@ 36,66 DCPRINT SAY NEWUPS->LREBDESC1
			@ 36,83 DCPRINT SAY NEWUPS->LREBCODE1
			@ 36,90 DCPRINT SAY NEWUPS->LREBATE1 PICTURE '$99,999.99'
		ENDIF

		IF NEWUPS->REBATE2 > 0
			@ 37,1 DCPRINT SAY NEWUPS->REBDESC2
			@ 37,20 DCPRINT SAY NEWUPS->REBCODE2
			@ 37,30 DCPRINT SAY NEWUPS->REBATE2 PICTURE '$99,999.99'
		ENDIF

		IF NEWUPS->LREBATE2 > 0
			@ 37,66 DCPRINT SAY NEWUPS->LREBDESC2
			@ 37,83 DCPRINT SAY NEWUPS->LREBCODE2
			@ 37,90 DCPRINT SAY NEWUPS->LREBATE2 PICTURE '$99,999.99'
		ENDIF

		IF NEWUPS->REBATE3 > 0
			@ 38,1 DCPRINT SAY NEWUPS->REBDESC3
			@ 38,20 DCPRINT SAY NEWUPS->REBCODE3
			@ 38,30 DCPRINT SAY NEWUPS->REBATE3 PICTURE '$99,999.99'
		ENDIF

		IF NEWUPS->LREBATE3 > 0
			@ 38,66 DCPRINT SAY NEWUPS->LREBDESC3
			@ 38,83 DCPRINT SAY NEWUPS->LREBCODE3
			@ 38,90 DCPRINT SAY NEWUPS->LREBATE3 PICTURE '$99,999.99'
		ENDIF

		IF NEWUPS->REBATE4 > 0
			@ 39,1 DCPRINT SAY NEWUPS->REBDESC4
			@ 39,20 DCPRINT SAY NEWUPS->REBCODE4
			@ 39,30 DCPRINT SAY NEWUPS->REBATE4 PICTURE '$99,999.99'
		ENDIF

		IF NEWUPS->LREBATE4 > 0
			@ 39,66 DCPRINT SAY NEWUPS->LREBDESC4
			@ 39,83 DCPRINT SAY NEWUPS->LREBCODE4
			@ 39,90 DCPRINT SAY NEWUPS->LREBATE4 PICTURE '$99,999.99'
		ENDIF

		IF NEWUPS->REBATE5 > 0
			@ 40,1 DCPRINT SAY NEWUPS->REBDESC5
			@ 40,20 DCPRINT SAY NEWUPS->REBCODE5
			@ 40,30 DCPRINT SAY NEWUPS->REBATE5 PICTURE '$99,999.99'
		ENDIF

		IF NEWUPS->LREBATE5 > 0
			@ 40,66 DCPRINT SAY NEWUPS->LREBDESC5
			@ 40,83 DCPRINT SAY NEWUPS->LREBCODE5
			@ 40,90 DCPRINT SAY NEWUPS->LREBATE5 PICTURE '$99,999.99'
		ENDIF

		IF NEWUPS->REBATE6 > 0
			@ 41,1 DCPRINT SAY NEWUPS->REBDESC6
			@ 41,20 DCPRINT SAY NEWUPS->REBCODE6
			@ 41,30 DCPRINT SAY NEWUPS->REBATE6 PICTURE '$99,999.99'
		ENDIF

		IF NEWUPS->LREBATE6 > 0
			@ 41,66 DCPRINT SAY NEWUPS->LREBDESC6
			@ 41,83 DCPRINT SAY NEWUPS->LREBCODE6
			@ 41,90 DCPRINT SAY NEWUPS->LREBATE6 PICTURE '$99,999.99'
		ENDIF



		@ 42,1 DCPRINT SAY 'Retail Rebates    '
		@ 42,30 DCPRINT SAY NEWUPS->REBATE  PICTURE '$99,999.99'
		@ 42,66 DCPRINT SAY 'Lease Rebates'
		@ 42,90 DCPRINT SAY NEWUPS->LREBATE PICTURE '$99,999.99'

		@ 43,0,53.2,64 DCPRINT BOX
		@ 43,65,53.2,130 DCPRINT BOX


	   IF lSideBySide
			@ 44,1 DCPRINT SAY 'Net Financed or Due'
		   @ 44,30 DCPRINT SAY (NEWUPS->QUOTED+NEWUPS->TAX+NEWUPS->AFTSALETAX+NEWUPS->LIEN)-(NEWUPS->REBATE+NEWUPS->DOWNPAY+NEWUPS->TRADEQUOTE)  picture '$999,999.99'

			IF lPrintRange //.OR. lPrintPayWords
			 @ 45,1 DCPRINT SAY 'Retail Financing'           FONT '10.Courier New Bold'
			 @ 46,1 DCPRINT SAY 'Term / Payment Range   '    FONT '10.Courier New Bold'
			 //@ 47,10 DCPRINT SAY '24 Month-> '+ cRange24Desc
			 @ 47,10 DCPRINT SAY '36 Month-> '+ cRange36Desc
			 @ 48,10 DCPRINT SAY '48 Month-> '+ cRange48Desc
			 @ 49,10 DCPRINT SAY '60 Month-> '+ cRange60Desc
			 @ 50,10 DCPRINT SAY '72 Month-> '+ cRange72Desc
			 @ 51,9 DCPRINT SAY str(NEWUPS->termodd)+' Month-> '+cRangeOddDesc
			else
			 @ 45,1 DCPRINT SAY 'Retail Financing'           FONT '10.Courier New Bold'
			 @ 46,1 DCPRINT SAY 'Term / Estimated Payment '    FONT '10.Courier New Bold'
			 //@ 47,10 DCPRINT SAY '24 Month-> '+ str(ROUND(NEWUPS->PAY24*nWalkbump,2)) FONT '10.Courier New Bold'
			 @ 47,10 DCPRINT SAY '36 Month-> '+ str(ROUND(NEWUPS->PAY36*nWalkbump,2)) FONT '10.Courier New Bold'
			 @ 48,10 DCPRINT SAY '48 Month-> '+ str(ROUND(NEWUPS->PAY48*nWalkbump,2)) FONT '10.Courier New Bold'
			 @ 49,10 DCPRINT SAY '60 Month-> '+ str(ROUND(NEWUPS->PAY60*nWalkbump,2)) FONT '10.Courier New Bold'
			 @ 50,10 DCPRINT SAY '72 Month-> '+ str(ROUND(NEWUPS->PAY72*nWalkbump,2)) FONT '10.Courier New Bold'
			 @ 51,9 DCPRINT SAY str(NEWUPS->termodd)+' Month-> '+str(ROUND(NEWUPS->PAYODD*nWalkbump,2)) FONT '10.Courier New Bold'
         ENDIF

		 IF NEWUPS->LPAYMENT2 > 0
			@ 44,66 DCPRINT SAY 'Capitalized Cost'
			@ 44,100 DCPRINT SAY NEWUPS->CAPCOST  PICTURE '$99,999.99'
			@ 45,66 DCPRINT SAY 'Leasing Payment'       FONT '10.Courier New Bold'
			IF ALLTRIM(NEWUPS->EXTRACHAR3)='Y'
				IF lShowPropTax
				 @ 45,90 DCPRINT SAY 'PropTax Included=$ '+ alltrim(STR(NEWUPS->EXTRANUM3)) FONT '10.Courier New Bold'
				else
				 @ 45,90 DCPRINT SAY 'PropTax Included'
				endif
			ENDIF
			@ 46,66 DCPRINT SAY 'Lease Term In Months '
			@ 46,100 DCPRINT SAY NEWUPS->LTERM         PICTURE '999'
			@ 47,66 DCPRINT SAY 'Lease Miles/Year'
			@ 47,100 DCPRINT SAY NEWUPS->LMILES+NEWUPS->XCESSMILE   PICTURE '99999'
			@ 48,66 DCPRINT SAY 'Est. Monthly Payment'
			@ 48,100 DCPRINT SAY NEWUPS->LPAYMENT2      PICTURE '$9,999.99'
			@ 49,66 DCPRINT SAY 'Residual'
			@ 49,100 DCPRINT SAY NEWUPS->RESIDUAL PICTURE '$99,999.99'
			@ 50,66 DCPRINT SAY 'Due At Sign'
			@ 50,100 DCPRINT SAY NEWUPS->DUEATSIGN PICTURE '$99,999.99'
			@ 51,66 DCPRINT SAY 'Taxes On Lease '
		   @ 51,100 DCPRINT SAY NEWUPS->leasetax picture '$99,999.99'
		   @ 52,66 DCPRINT SAY 'Finance/Lease Company '
		   @ 52,100 DCPRINT SAY NEWUPS->lCompany
		 ENDIF
		endif        // SIDEBYSIDE TRUE
		IF !lSideBySide
			IF NEWUPS->QUOTETYPE='R'
			 @ 44,1 DCPRINT SAY 'Net Financed or Due'
		    @ 44,30 DCPRINT SAY (NEWUPS->QUOTED+NEWUPS->TAX+NEWUPS->AFTSALETAX+NEWUPS->LIEN)-(NEWUPS->REBATE+NEWUPS->DOWNPAY+NEWUPS->TRADEQUOTE)  picture '$999,999.99'
			 IF lPrintRange
			  @ 46,1 DCPRINT SAY 'Term / Payment Range   $'    FONT '10.Courier New Bold'
			 //@ 47,10 DCPRINT SAY '24 Month-> '+ cRange24Desc
			  @ 47,10 DCPRINT SAY '36 Month-> '+ cRange36Desc
			  @ 48,10 DCPRINT SAY '48 Month-> '+ cRange48Desc
			  @ 49,10 DCPRINT SAY '60 Month-> '+cRange60Desc
			  @ 50,10 DCPRINT SAY '72 Month-> '+ cRange72Desc

			  @ 51,9 DCPRINT SAY str(NEWUPS->termodd)+' Month-> '+cRangeOddDesc ///STR(ROUND(NEWUPS->payodd,0))+' '+str(ROUND(NEWUPS->payodd*nWalkbump,0)) FONT '10.Courier New Bold'
			else
			  @ 46,1 DCPRINT SAY 'Term / Estimated Payment '    FONT '10.Courier New Bold'
			 //@ 47,10 DCPRINT SAY '24 Month-> '+ str(ROUND(NEWUPS->PAY24*nWalkbump,2)) FONT '10.Courier New Bold'
			  @ 47,10 DCPRINT SAY '36 Month-> $'+ str(ROUND(NEWUPS->PAY36*nWalkbump,2)) FONT '10.Courier New Bold'
			  @ 48,10 DCPRINT SAY '48 Month-> $'+ str(ROUND(NEWUPS->PAY48*nWalkbump,2)) FONT '10.Courier New Bold'
			  @ 49,10 DCPRINT SAY '60 Month-> $'+ str(ROUND(NEWUPS->PAY60*nWalkbump,2)) FONT '10.Courier New Bold'
			  @ 50,10 DCPRINT SAY '72 Month-> $'+ str(ROUND(NEWUPS->PAY72*nWalkbump,2)) FONT '10.Courier New Bold'

			  @ 51,9 DCPRINT SAY str(NEWUPS->termodd)+' Month-> $ '+str(ROUND(NEWUPS->PAYODD*nWalkbump,2)) FONT '10.Courier New Bold'
         ENDIF
		  ENDIF
		  IF NEWUPS->QUOTETYPE='L'
		   IF NEWUPS->LPAYMENT2 > 0
				@ 44,66 DCPRINT SAY 'Capitalized Cost'
			   @ 44,100 DCPRINT SAY NEWUPS->CAPCOST     PICTURE '$99,999.99'
			   @ 45,66 DCPRINT SAY 'Leasing Payment'           FONT '10.Courier New Bold'
				IF ALLTRIM(NEWUPS->EXTRACHAR3)='Y'
				 IF lShowPropTax
				  @ 45,90 DCPRINT SAY 'PropTax Included=$ '+ alltrim(STR(NEWUPS->EXTRANUM3)) FONT '10.Courier New Bold'
				 else
				  @ 45,90 DCPRINT SAY 'PropTax Included'
				 endif
			   ENDIF
			   @ 46,66 DCPRINT SAY 'Lease Term In Months '
			   @ 46,100 DCPRINT SAY NEWUPS->LTERM         PICTURE '999'
			   @ 47,66 DCPRINT SAY 'Est. Monthly Payment'
			   @ 47,100 DCPRINT SAY NEWUPS->LPAYMENT2      PICTURE '$9,999.99'
			   @ 48,66 DCPRINT SAY 'Residual'
				@ 48,100 DCPRINT SAY NEWUPS->RESIDUAL PICTURE '$99,999.99'
				@ 49,66 DCPRINT SAY 'Due At Sign+DMV'
				@ 49,100 DCPRINT SAY NEWUPS->DUEATSIGN PICTURE '$99,999.99'
				@ 50,66 DCPRINT SAY 'Taxes On Lease '
		   	@ 50,100 DCPRINT SAY NEWUPS->leasetax picture '$99,999.99'
		   	@ 51,66 DCPRINT SAY 'Finance/Lease Company '
		   	@ 51,100 DCPRINT SAY NEWUPS->lCompany
		  ENDIF
		 ENDIF
		ENDIF        /// NOT SIDEBYSIDE

		@ 53,0 DCPRINT SAY 'All quoted payments are ESTIMATES. Interest rates are subject to change ' font '10.Courier New Bold'
		@ dc_printerrow()+1,0 DCPRINT SAY 'based on manufacturer program changes, expirations and   ' font '10.Courier New Bold'
		@ dc_printerrow()+1,0 DCPRINT SAY 'lending banks interest rates and credit decisions.'   font '10.Courier New Bold'
		@ dc_printerrow()+2,0 DCPRINT SAY ' No Monthly Payment is certain until lending bank confirmation.' font '12.Courier New Bold'
		@ dc_printerrow()+2,0 DCPRINT SAY "THIS QUOTE IS NOT VALID WITHOUT A MANAGER'S SIGNATURE"
		@ dc_printerrow()+2,0 DCPRINT SAY 'SALESPERSON NUMBER AND SIGNATURE '+'     '+NEWUPS->SALESMAN
		@ dc_printerrow(),DC_PRINTERCOL()+15  DCPRINT SAY '__________________________________ '
		@ DC_PRINTERROW()+2,0 DCPRINT SAY 'MANAGER ' + NEWUPS->TOMGR +' '+'_____________________________________'
		nCopies:=nCopies-1
		IF nCopies > 0
			DCPRINT EJECT
		ENDIF
	ENDDO
	DCPRINT OFF PRINTER OPRINTER
RETURN NIL

static function _getRangeDescX(nRange,cRangeDesc,nTerm)
	 local nBumpMinus:=20,nBumpPlus:=35

	 IF nRange < 20
		cRangeDesc:='0'
		RETURN cRangedesc
	 ENDIF
	 IF nTerm <= 84
		  nBumpMinus:=20
		  nBumpPlus:=40
	 ENDIF

	 IF nTerm < 61
		  nBumpMinus:=20
		  nBumpPlus:=35
	 ENDIF

	 IF nTerm < 49
		  nBumpMinus:=15
		  nBumpPlus:=30
	 ENDIF

	 IF nTerm < 36
		  nBumpMinus:=10
		  nBumpPlus:=25
	 ENDIF


	 IF nTerm < 25
		  nBumpMinus:=5
		  nBumpPlus:=20
	 ENDIF

	 cRangeDesc:=alltrim(str(round(nRange-nBumpMinus,2)))+' to '+alltrim(str(round(nRange+nBumpPlus,2)))
	 return cRangedesc

static function _getRangeDesc(nRange,cRangeDesc)
	nRange:=round(nRange,0)

  IF nRange =0
		cRangeDesc:='0'
		RETURN cRangedesc
  ENDIF


  IF nRange > 0
	IF nRange <= 49
		 cRangeDesc:='Under $50'
	ENDIF
  ENDIF
  IF nRange >= 50
	IF nRange <= 99
		 cRangeDesc:='Under $100'
	ENDIF
  ENDIF

  IF nRange >=100
	IF nRange <= 135
		 cRangeDesc:='Low One Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 136
	IF nRange <= 165
		 cRangeDesc:='Middle One Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 166
	IF nRange <= 199
		 cRangeDesc:='High One Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >=200
	IF nRange <= 235
		 cRangeDesc:='Low Two Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 236
	IF nRange <= 265
		 cRangeDesc:='Middle Two Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 266
	IF nRange <= 299
		 cRangeDesc:='High Two Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >=300
	IF nRange <= 335
		 cRangeDesc:='Low Three Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 336
	IF nRange <= 365
		 cRangeDesc:='Middle Three Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 366
	IF nRange <= 399
		 cRangeDesc:='High Three Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >=400
	IF nRange <= 435
		 cRangeDesc:='Low Four Hundred $$ Range '
	ENDIF
  ENDIF
  IF nRange >= 436
	IF nRange <= 465
		 cRangeDesc:='Middle Four Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 466
	IF nRange <= 499
		 cRangeDesc:='High Four Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >=500
	IF nRange <= 535
		 cRangeDesc:='Low Five Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 536
	IF nRange <= 565
		 cRangeDesc:='Middle Five Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 566
	IF nRange <= 599
		 cRangeDesc:='High Five Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >=600
	IF nRange <= 635
		 cRangeDesc:='Low Six Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 636
	IF nRange <= 665
		 cRangeDesc:='Middle Six Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 666
	IF nRange <= 699
		 cRangeDesc:='High Six Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >=700
	IF nRange <= 735
		 cRangeDesc:='Low Seven Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 736
	IF nRange <= 765
		 cRangeDesc:='Middle Seven Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 766
	IF nRange <= 799
		 cRangeDesc:='High Seven Hundred $$ Range'
	ENDIF
  ENDIF

  IF nRange >=800
	IF nRange <= 835
		 cRangeDesc:='Low Eight Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 836
	IF nRange <= 865
		 cRangeDesc:='Middle Eight Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 866
	IF nRange <= 899
		 cRangeDesc:='High Eight Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >=900
	IF nRange <= 935
		 cRangeDesc:='Low Nine Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 936
	IF nRange <= 965
		 cRangeDesc:='Middle Nine Hundred $$ Range'
	ENDIF
  ENDIF
  IF nRange >= 966
	IF nRange <= 999
		 cRangeDesc:='High Nine Hundred $$ Range'
	ENDIF
  ENDIF

  return cRangeDesc


static function savaquote(cBody,oRecord)
	////cVehicleID,nQuotePrice,nTissue,nSalesTax,nTaxrate,nTradeallow,nTradeACV)
	local clSavequote:='Y'
	local getlist:={} ,nRec
	local GETOPTIONS, xStatus
	DCGETOPTIONS SAYWIDTH 0 SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE TABSTOP


	//// LETS STOP SAVING THIS FOR NOW to see if the bloat ptoblem is here   BV 02/04/2017


	@ 5,0 DCSAY 'Attention If you would like to save this latest quote to the comment record enter Y.'
	@ 6,0 DCSAY 'If nothing has changed you should enter N'
	@ 8,50 DCSAY 'Save this Quote Y/N' get clSavequote PICTURE '@! L'
	DCREAD GUI MODAL TO xStatus ENTEREXIT FIT OPTIONS GETOPTIONS TITLE ' ' EVAL{|o| SETAPPWINDOW(o)}
	if !xStatus
	    RETURN oRecord
	endif
	IF clSavequote='N'
		return oRecord
	ENDIF

	cBody:='.'+CRLF
	cBody:=cBody+'..'+CRLF
	cBody:=cBody+'Quote....'+ CRLF
	cBody:=cBody+dtos(date())+ ' '+time()+ CRLF
	cBody:=cBody+'Stock Number '+ oRecord:stkquoted +' '+'Last6  '+SUBSTR(oRecord:QUOTEVIN,12,6)+CRLF
	cBody:=cBody+oRecord:yrquoted+' '+oRecord:vehquoted +CRLF
	cBody:=cBody+'Ordertag '+ str((oRecord:Quoted-oRecord:QuoteTiss)-(oRecord:Tradequote-oRecord:TradeACV))+CRLF
	IF oRecord:quotetype $('CR')
		cBody:=cBody+ 'Your Price ' + str(oRecord:quoted) +CRLF
		cBody:=cBody+ 'Sales tax  ' + str(round((oRecord:Tax+oRecord:aftsaletax),2))+' '+'TaxRate '+ str(oRecord:Taxper) +CRLF
		cBody:=cBody+ 'PreEstimated Fees ' + alltrim(str(oRecord:regfee+oRecord:procfee+oRecord:wastefee+oRecord:inspection))+CRLF

		IF oRecord:rebate1 > 0
			cBody:=cBody+ oRecord:Rebdesc1+'  ' + alltrim(str(oRecord:Rebate1)) +CRLF
		ENDIF
		IF oRecord:rebate2 > 0
			cBody:=cBody+ oRecord:Rebdesc2+'  ' + alltrim(str(oRecord:Rebate2)) +CRLF
		ENDIF
		IF oRecord:rebate3 > 0
			cBody:=cBody+ oRecord:Rebdesc3+'  ' + alltrim(str(oRecord:Rebate3)) +CRLF
		ENDIF
		IF oRecord:rebate4 > 0
			cBody:=cBody+ oRecord:Rebdesc4+'  ' + alltrim(str(oRecord:Rebate4)) +CRLF
		ENDIF
		IF oRecord:rebate5 > 0
			cBody:=cBody+ oRecord:Rebdesc5+'  ' + alltrim(str(oRecord:Rebate5)) +CRLF
		ENDIF
		IF oRecord:rebate6 > 0
			cBody:=cBody+ oRecord:Rebdesc6+'  ' + alltrim(str(oRecord:Rebate6)) +CRLF
		ENDIF
		cBody:=cBody+ 'Total Rebates   ' + str(oRecord:rebate) +CRLF
		cBody:=cBody+ 'Net Price  ' + str((oRecord:QUOTED+oRecord:Tax)-oRecord:REBATE)  +CRLF
	ENDIF
	cBody:=cBody+ 'TRADE IN INFORMATION:'+ ' ' + oRecord:TRADEYR+' '+oRecord:TRADEmake+' '+oRecord:TRADEDESC +CRLF
	cBody:=cBody+ 'TRADE Allowance     :'+ ' ' + str(oRecord:TRADEquote) +CRLF
	cBody:=cBody+ 'TRADE LienPayoff    :'+ ' ' + str(oRecord:Lien)+CRLF
	cBody:=cBody+ 'TRADE ACV           :'+ ' ' + str(oRecord:TRADEACV)+CRLF


	IF oRecord:QUOTETYPE='R'
		cBody:=cBody+  'FINANCE PAYMENT QUOTES:' +CRLF
		cBody:=cBody+  'AMOUNT FINANCED ' + str(round(oRecord:FINANCE,2))+CRLF
		cBody:=cBody+  'DOWNPAYMENT' + str(oRecord:DOWNPAY) +CRLF
		cBody:=cBody+  'TERM / EST. PAYMENT / RATE' + CRLF
		cBody:=cBody+  '36MO$ '+STR(oRecord:PAY36)+'  '+str(oRecord:Rate36)+CRLF
		cBody:=cBody+  '48MO$ '+STR(oRecord:PAY48)+'  '+str(oRecord:Rate48)+CRLF
		cBody:=cBody+  '60MO$ '+STR(oRecord:PAY60)+'  '+str(oRecord:Rate60)+CRLF
		cBody:=cBody+  '72MO$ '+STR(oRecord:PAY72)+'  '+str(oRecord:Rate72)+CRLF
		cBody:=cBody+  str(oRecord:termodd)+'$'+STR(oRecord:payodd)+'  '+str(oRecord:intrate)+CRLF
	endif
	IF NEWUPS->QUOTETYPE='L'
		cBody:=cBody+ 'LEASE TERM IN MONTHS  ' +str(oRecord:LTERM) +  ' LEASE EST. MONTHLY PAYMENT  $ '+ str(oRecord:LPAYMENT );
			+'  Residual  $ '+str(oRecord:residual)+CRLF
	ENDIF
	// update record object'

	oRecord:Notes:=alltrim(oRecord:Notes)+ALLTRIM(cBody)

	IF len(oRecord:notes) > 500000
		oRecord:notes:=""
	ENDIF
return oRecord

static function saveaNapp(cTracker,cBody)

	cBody:='.'+CRLF
	cBody:=cBody+'..'+CRLF
	cBody:=cBody+'Credit App....'+cTracker+ CRLF
	cBody:=cBody+dtos(date())+ ' '+time()+ CRLF
	cBody:=cBody+'Stock Number '+ acevehic->stocnum +' '+'vin  '+acevehic->vin +CRLF
	cBody:=cBody+acevehic->year+' '+acevehic->bodystyle +CRLF
	cBody:=cBody+'Ordertag '+ str(acedeal->price-acevehic->cost) +CRLF
	cBody:=cBody+ 'Your Price ' + str(acedeal->price) +CRLF
	cBody:=cBody+ 'Sales tax  ' + str(acedeal->salestax)+' '+'TaxRate '+ str(acedeal->salestaxpr) +CRLF
	cBody:=cBody+ 'Rebate     ' + str(acedeal->rebate) +CRLF
	cBody:=cBody+ 'Net Price  ' + str((acedeal->Price+Acedeal->salestax)-acedeal->REBATE)  +CRLF
	IF acedeal->nFintype <> 2
		cBody:=cBody+  'FINANCE PAYMENT QUOTES:' +CRLF
		cBody:=cBody+  'AMOUNT FINANCED ' + str(acedeal->netamtfin)+CRLF
		cBody:=cBody+  'Interest rate '+ str(acedeal->aprrtlgkp) +CRLF
		cBody:=cBody+  'DOWNPAYMENT' + str(acedeal->cashdown) +CRLF
		cBody:=cBody+  'TERM ' + str(acedeal->term)+CRLF
		cBody:=cBody+  'Payment '+STR(acedeal->monthlypay) +CRLF
	endif
	IF acedeal->nFintype=2
		cBody:=cBody+ 'LEASE TERM IN MONTHS  ' +str(acedeal->term) +  '  Residual  $ '+str(acedeal->resamount)+CRLF
		cBody:=cBody+ 'Lease Interest Rate '+str(acedeal->leaseper) +CRLF
		cBody:=cBody+ 'Lease Money Factor  '+str(acedeal->leasefact) +CRLF
		cBody:=cBody+ 'MSRP GUide          '+str(acedeal->nMSRPguide)+CRLF
		cBody:=cBody+ 'MSRP adds           '+str(acedeal->nMSRPadd) +CRLF
		cBody:=cBody+ 'Residual  $         '+str(acedeal->resamount)+CRLF
		cbody:=cBody+ 'Total Cap Cost      '+str(acedeal->totcapcost)+CRLF
		cbody:=cBody+ 'EST. Lease Payment '+str(acedeal->totleaspay)+CRLF
		cBody:=cBody+ 'Payment with Property Tax' +str(Acedeal->paywproptx)+CRLF
	ENDIF

	cBody:=cBody+ 'TRADE IN INFORMATION:'+ ' ' + acetrade->TRyear+' '+acetrade->TRmake+' '+acetrade->trmodel+CRLF
	cBody:=cBody+ 'TRADE Allowance     :'+ ' ' + str(acedeal->TRADEallow) +CRLF
	cBody:=cBody+ 'TRADE LienPayoff    :'+ ' ' + str(acedeal->oweontrade)+CRLF


return nil


FUNCTION CHKSTKNUM(oRecord,GETLIST,lchk4chg)
	local nChoice:=0
	default lchk4chg:=.T.

	IF EMPTY(oRecord:NEWUSED)
		DC_WINALERT('Please Select New Or Used')
		RETURN TRUE
	ENDIF

	IF !lchk4chg
		IF EMPTY(NEWUPS->STKQUOTED)
		  RETURN TRUE
		ENDIF
	ENDIF

	IF UPPER(oRecord:STKQUOTED)='LOCATE' .OR. UPPER(oRecord:STKQUOTED)='ORDER' .OR. UPPER(oRecord:STKQUOTED)='LBO'
		IF oRecord:QUOTEMSRP=0
		   oRecord:QUOTETISS:=oRecord:QUOTEMSRP:=0
		ENDIF
		IF EMPTY(oRecord:QUOTEVIN)
		 oRecord:QUOTEVIN := SPACE(17)
		ENDIF
		IF EMPTY(oRecord:YRQUOTED)
		 oRecord:YRQUOTED:=SPACE(4)
		ENDIF
		IF EMPTY(oRecord:VEHQUOTED)
		 oRecord:VEHQUOTED:=SPACE(20)
		ENDIF
		DC_GETREFRESH(GETLIST)
		RETURN .T.
	ENDIF


	IF oRecord:STKQUOTED <> NEWUPS->STKQUOTED
	   NCHOICE:=confirmbox(,'You Have Changed Stock Numbers- enter Yes to Zero all Calculations ','Zero Calculations ?',XBPMB_YESNO,XBPMB_QUESTION+;
			XBPMB_APPMODAL+XBPMB_MOVEABLE)
		IF NCHOICE=6
			oRecord:QUOTED:=oRecord:FINANCE:=oRecord:PAY72:=oRecord:PAY36:=oRecord:PAY48:=;
			oRecord:PAY60:=;
         oRecord:TERMODD:=oRecord:PAYODD:=oRecord:TERM:=oRecord:LPAYMENT:=oRecord:LTERM2:=;
		   oRecord:LPAYMENT2:=oRecord:RESIDUAL:=oRecord:RESIDUAL2:=oRecord:REBATE1:=oRecord:REBATE2:=;
			oRecord:REBATE3:=oRecord:REBATE4:=oRecord:REBATE5:=oRecord:REBATE6:=oRecord:REBATE:=0
		   oRecord:REBDESC1:=oRecord:REBDESC2:=oRecord:REBDESC3:=oRecord:REBDESC4:=oRecord:REBDESC5:=;
			oRecord:REBDESC6:=space(10)
			oRecord:REBCODE1:=oRecord:REBCODE2:=oRecord:REBCODE3:=oRecord:REBCODE4:=oRecord:REBCODE5:=;
			oRecord:REBCODE6:=space(10)
			oRecord:LREBATE1:=oRecord:LREBATE2:=;
			oRecord:LREBATE3:=oRecord:LREBATE4:=oRecord:LREBATE5:=oRecord:LREBATE6:=oRecord:LREBATE:=0
		   oRecord:LREBDESC1:=oRecord:LREBDESC2:=oRecord:LREBDESC3:=oRecord:LREBDESC4:=oRecord:LREBDESC5:=;
			oRecord:LREBDESC6:=space(10)
			oRecord:LREBCODE1:=oRecord:LREBCODE2:=oRecord:LREBCODE3:=oRecord:LREBCODE4:=oRecord:LREBCODE5:=;
			oRecord:LREBCODE6:=space(10)
		ENDIF
	ENDIF


	IF UPPER(oRecord:NEWUSED)='N'
		SELECT NCI
		ORDSETFOCUS('INVIDX01')
		SEEK oRecord:STKQUOTED
		IF FOUND()
			oRecord:YRQUOTED:=NCI->YEAR
			oRecord:VEHQUOTED:=ALLTRIM(NCI->MAKE)+' '+ALLTRIM(NCI->MODEL)+' '+ALLTRIM(NCI->TRIM)+' '+ALLTRIM(NCI->COLOR)
			oRecord:QUOTETISS:=NCI->FLR_COST
			oRecord:QUOTEVIN:=NCI->VIN
			oRecord:QUOTEMSRP:=NCI->RETAIL
			IF oRecord:QUOTED < 1
				IF NCI->EPRICE > 0
				 oRecord:QUOTED :=NCI->EPRICE
		   	ELSE
				 oRecord:QUOTED :=NCI->RETAIL
				ENDIF
			ENDIF

			//IF oRecord:QUOTEMILES=0
				oRecord:QUOTEMILES:=NCI->MILEAGE
			//ENDIF
			oRecord:VEHQUOTED:=oRecord:VEHQUOTED+SPACE(20-LEN(oRecord:VEHQUOTED))
		   DC_GETREFRESH(GETLIST)
		   RETURN TRUE
		else
			Vehbrow(oRecord:NEWUSED,@oRecord:STKQUOTED)
		ENDIF
		getNEWschedcost(oRecord:STKQUOTED,@oRecord:QUOTETISS)

	ENDIF
	IF UPPER(oRecord:NEWUSED)='U'
		SELECT UCL01
		ORDSETFOCUS('INVIDX01')
		SEEK oRecord:STKQUOTED
		IF FOUND()
			oRecord:YRQUOTED:=UCL01->YEAR
			oRecord:VEHQUOTED:=ALLTRIM(UCL01->MAKE)+' '+ALLTRIM(UCL01->MODEL)+' '+ALLTRIM(UCL01->COLOR)
			oRecord:QUOTEVIN:=UCL01->VIN
			oRecord:QUOTEMSRP:=UCL01->retail
			IF oRecord:QUOTED < 1
				IF UCL01->EPRICE > 0
					oRecord:QUOTED :=UCL01->EPRICE
				ELSE
					oRecord:QUOTED :=UCL01->RETAIL
				ENDIF
			ENDIF
			//IF oRecord:QUOTEMILES=0
				oRecord:QUOTEMILES:=UCL01->MILEAGE
			//ENDIF
			ELSE
			Vehbrow(oRecord:NEWUSED,@oRecord:STKQUOTED)
		ENDIF
		GETSCHEDCOST(oRecord:STKQUOTED,@oRecord:QUOTETISS)
		oRecord:VEHQUOTED:=oRecord:VEHQUOTED+SPACE(20-LEN(oRecord:VEHQUOTED))
		DC_GETREFRESH(GETLIST)
		RETURN TRUE
	ENDIF
	DC_GETREFRESH(GETLIST)
RETURN FALSE

FUNCTION ASLK(cAdSource)
	LOCAL GetList := {}, aPres, MoBrowse, MoToolBar
	LOCAL GETOPTIONS , xStatus
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE
	SELECT adsource
	GOTO TOP
	IF !empty(cAdSource)
	  SEEK cAdsource
	  IF found()
		RETURN TRUE
	  ENDIF
	ENDIF
	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, -1 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 10 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

	@ 1,1 DCTOOLBAR MoToolBar                               ;
		SIZE 90, 1.5
/* ----- Create browse ----- */
	@ 3,1 DCBROWSE MoBrowse ALIAS 'Adsource'                 ;
		SIZE 50,25                                     ;
		FONT '10.Courier New'  ;
		PRESENTATION aPres ;
		datalink{|| loadasrce(@cAdsource), DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}

	DCBROWSECOL FIELD adsource->adsrce                     ;
		WIDTH 20 HEADER "Source" PARENT MoBrowse
	DCBROWSECOL FIELD adsource->desc                     ;
		WIDTH 20 HEADER "Desc" PARENT MoBrowse

	DCREAD GUI ;
		TO XSTATUS;
		EVAL {||SETAPPFOCUS(MobROWSE:GETCOLUMN(1))};
		FIT TITLE 'Up Ad Source Browser' ;
		OPTIONS GETOPTIONS;
		BUTTONS DCGUI_BUTTON_OK      ;
		APPWINDOW MAINWINDOW():DRAWINGAREA

	IF !XSTATUS
		RETURN TRUE
	ENDIF

	cAdsource:=adsource->adsrce
RETURN .t.



static function loadasrce(as)
	as:=adsource->adsrce
return .T.


FUNCTION SEEKZIP(oRecord)                       // PROGRAM TO FIND ZIP
	SELECT USAZIP
	USAZIP->(ORDSETFOCUS('USAZIP'))
	SEEK oRecord:ZIPCODE
	IF FOUND()
		oRecord:CITYSTATE:=USAZIP->CITY
		oRecord:STATE:=USAZIP->STATE
		oRecord:COUNTY:=USAZIP->COUNTY
	ENDIF
RETURN oRecord





* PROGRAM TO DISPLAY INTEREST AT UPENTRY
FUNCTION NPRODLK(cVehInterest,cFranchise)
	LOCAL GetList := {}, aPres, MoBrowse, MoToolBar , xStatus
	LOCAL GETOPTIONS
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE
	SELECT CLASSORT
	GOTO TOP

	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, -1 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 10 }  }              /* Cell Height */


/* ----- Create browse ----- */
	@ 3,1 DCBROWSE MoBrowse ALIAS 'CLASSORT'                 ;
		SIZE 50,25                                     ;
		FONT '10.Courier New'  ;
		PRESENTATION aPres ;
		datalink{|| loadnprod(@cVehInterest,@cFranchise), DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)};
		FREEZELEFT{1,2,3}

	DCBROWSECOL FIELD classort->CLASSTYPE                     ;
		WIDTH 20 HEADER "Interest" PARENT MoBrowse
	DCBROWSECOL FIELD classort->Franchise                     ;
		WIDTH 20 HEADER "Franchise" PARENT MoBrowse

	DCREAD GUI ;
		TO XSTATUS;
		EVAL {||SETAPPFOCUS(MobROWSE:GETCOLUMN(1))};
		FIT TITLE 'Up Interest Browser' ;
		OPTIONS GETOPTIONS;
		BUTTONS DCGUI_BUTTON_OK      ;
		APPWINDOW MAINWINDOW():DRAWINGAREA
	IF !XSTATUS
		RETURN .T.
	ENDIF
	cVehinterest:=classort->classtype
	cFranchise:=classort->franchise
RETURN .t.



* -------------

function BrowNProd()

   LOCAL GetList[0], GetOptions
   LOCAL aPres
   LOCAL oBrowse, oToolBar
   LOCAL lStatus

	IF !UseDb({'CLASSORT'})
     ALERTWIN TEXT "Cannot Open Files" TITLE "Data Alert" TIMEOUT 5
     return nil
	ENDIF

   DbSelectArea('CLASSORT')
   CLASSSORT->(DbGoTop())

   @   1.00,  1.00 DCTOOLBAR oToolBar SIZE 60, 1.5

   DCADDBUTTON CAPTION "Add Interest Class";
      ACTION {|| AddInt(), DC_GetRefresh(GetList), oBrowse:forcestable()} ;
      PARENT oToolBar SIZE 30;
      TOOLTIP "Add Interest"

   DCADDBUTTON CAPTION "Delete Interest Class";
      ACTION {|| DelInt(classort->(RecNo())), DC_GetRefresh(GetList), oBrowse:forcestable()};
      PARENT oToolBar SIZE 30 ;
      TOOLTIP "Delete Interest"

   // CREATE BROWSE
   @   3.00,  1.00 DCBROWSE oBrowse ALIAS 'CLASSORT' SIZE 50,25;
                   FONT '10.Courier New'  ;
                   PRESENTATION ;
                               {{ XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },;             // Header FG Color
                                { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY },;          // Header BG Color
                                { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },;  // Row Sep
                                { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },;  // Col Sep
                                { XBP_PP_COL_DA_ROWHEIGHT, -1 },;                    // Row Height
                                { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY },;   // hilite BG Color
                                { XBP_PP_COL_DA_CELLHEIGHT, 10 }};                   // Cell Height
                   EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITEXIT


   DCBROWSECOL FIELD classort->CLASSTYPE;
      WIDTH 20 HEADER "Interest" PARENT oBrowse

   DCBROWSECOL FIELD classort->Franchise;
      WIDTH 20 HEADER "Franchise" PARENT oBrowse

   DCGETOPTIONS TITLE "Up Interest Browser" ;
     SAYWIDTH 0 ;
     AUTORESIZE

   DCREAD GUI TO lStatus OPTIONS GetOptions FIT BUTTONS DCGUI_BUTTON_OK AppWindow MainWindow():DrawingArea SETFOCUS @oBrowse

   IF !lStatus
     CLOSE ALL
     return .T.
	ENDIF
   classort->(DbCloseArea())

	close all
RETURN .t.

* -------------

function BrowZipper()

   LOCAL GetList[0], GetOptions
   LOCAL oBrowse, oToolBar
   LOCAL lStatus

   IF !UseDb({'ZIPPER'}, 1, .F., .T.)
     ALERTWIN TEXT "Cannot Open Files" TITLE "Data Alert" TIMEOUT 5
     return nil
	ENDIF

   @   1.00,  1.00 DCTOOLBAR oToolBar SIZE 90, 1.5

//   DCADDBUTTON CAPTION "Add Zip Code";
//      ACTION {|| zipadd(), DC_GetRefresh(GetList), moBrowse:forcestable()} ;
//      PARENT oToolBar SIZE 30 ;
//      TOOLTIP "Add Zip Code"

   DCADDBUTTON CAPTION "Add Zip Code";
      ACTION {|| /*ZipAdd(),*/ DC_GetRefresh(GetList), oBrowse:forcestable()};
      PARENT oToolBar SIZE 30;
      TOOLTIP "Add A Zip Code"

   DCADDBUTTON CAPTION "Delete Zip Code" ;
      ACTION {|| DelZip(zipper->(RecNo())), DC_GetRefresh(GetList), oBrowse:forcestable()};
      PARENT oToolBar SIZE 30 ;
      TOOLTIP "Delete Zip Code"

   DCADDBUTTON CAPTION "Print List" ;
      ACTION {|| ListZip(), DC_GetRefresh(GetList), oBrowse:forcestable()} ;
      PARENT oToolBar SIZE 30 ;
      TOOLTIP "Add Zip Code"

   DCADDBUTTON CAPTION "Print List";
      ACTION {|| ListZip(), DC_GetRefresh(GetList), oBrowse:forcestable()};
      PARENT oToolBar SIZE 30 ;
      TOOLTIP "Add Zip Code"

   @   3.00,  1.00 DCBROWSE oBrowse ALIAS 'ZIPPER' SIZE 80,25 ;
                   FONT '10.Courier New'  ;
                   PRESENTATION {{ XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },;
                                 { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY },;
                                 { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },;
                                 { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },;
                                 { XBP_PP_COL_DA_ROWHEIGHT, -1 },;
                                 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY },;
                                 { XBP_PP_COL_DA_CELLHEIGHT, 10 }};
                   SORTSCOLOR 0,2 ;
                   SORTUCOLOR 1,6   ;
                   EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITEXIT

   DCBROWSECOL FIELD zipper->zipcode ;
      WIDTH 9 HEADER "**Zipcode" PARENT oBrowse HCOLOR 1,5  ;
      SORT{|| ZIPPER->(OrdSetFocus('ZIPPERIN')), DbGoTop(), oBrowse:refreshAll()}

   DCBROWSECOL FIELD zipper->citystate;
      WIDTH 30 HEADER "**CityState" PARENT oBrowse HCOLOR 1,5 ;
      SORT{|| ZIPPER->(OrdSetFocus('ZIPALPHA')), DbGoTop(), oBrowse:refreshAll()}

   DCBROWSECOL FIELD zipper->adzone;
      WIDTH 5 HEADER "**AdZone" PARENT oBrowse HCOLOR 1,5 ;
      SORT{|| ZIPPER->(OrdSetFocus('ZIPZONE')), DbGoTop(), oBrowse:refreshAll()}

   DCBROWSECOL FIELD zipper->area ;
      WIDTH 5 HEADER "Areacode" PARENT oBrowse HCOLOR 1,5

   DCGETOPTIONS TITLE "Local ZipCode Browser";
     SAYWIDTH 0;
     AUTORESIZE

   DCREAD GUI TO lStatus OPTIONS GetOptions FIT BUTTONS DCGUI_BUTTON_OK APPWINDOW MainWindow():DrawingArea SETFOCUS @oBrowse

   IF !lStatus
     CLOSE ALL
     RETURN .T.
	ENDIF

   CLOSE ALL

return .t.

* -------------

FUNCTION BrowRebates()
	LOCAL GetList := {}, aPres, oBrowse, oToolBar
	LOCAL GETOPTIONS   , xStatus
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE


	IF UseDb({'REBATEUP'})
		REBATEUP->(DBGOTOP())
	ELSE
		DC_WINALERT('Cannot Open Files     .')
	    RETURN NIL
	ENDIF

	SET DELETED ON
	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },       ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY },    ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },         ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },         ;
    { XBP_PP_COL_DA_ROWHEIGHT, -1 },                        ;
    { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY },        ;
    { XBP_PP_COL_DA_CELLHEIGHT, 10 }  }

	@ 1,1 DCTOOLBAR oToolBar                               ;
		SIZE 90, 1.5

	DCADDBUTTON CAPTION 'Add Rebate Code'                              ;
		SIZE 30                                              ;
		ACTION {|| _REBATEADD(), DC_GetRefresh(GetList),obrowse:forcestable()}        ;
		PARENT oToolBar                                     ;
		TOOLTIP 'Add Zip Code'

	DCADDBUTTON CAPTION 'Delete Rebate Code'                              ;
		SIZE 30                                              ;
		ACTION {|| _REBATEDEL(REBATEUP->(recno())), DC_GetRefresh(GetList),obrowse:forcestable()}        ;
		PARENT oToolBar                                     ;
		TOOLTIP 'Delete Zip Code'

	DCADDBUTTON CAPTION 'Print List'                              ;
		SIZE 30                                              ;
		ACTION {|| _REBATELIST(), DC_GetRefresh(GetList),obrowse:forcestable()}        ;
		PARENT oToolBar                                     ;
		TOOLTIP 'Add Zip Code'



	@ 3,1 DCBROWSE oBrowse ALIAS 'REBATEUP'                 ;
		SIZE 90,25                                     ;
		FONT '10.Courier New'  ;
		PRESENTATION aPres     ;
		EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITEXIT


	DCBROWSECOL FIELD REBATEUP->REBCODE                     ;
		WIDTH 8 HEADER "Rebate Code" PARENT oBrowse  ;

	DCBROWSECOL FIELD REBATEUP->REBDESC                     ;
		WIDTH 30 HEADER "Description" PARENT oBrowse  ;

	DCBROWSECOL FIELD REBATEUP->AMOUNT                     ;
		WIDTH 10 HEADER "Amount" PARENT oBrowse PICTURE '99999.99' ;


	DCREAD GUI ;
		TO XSTATUS;
		FIT TITLE 'Rebate Browser' ;
		OPTIONS GETOPTIONS;
		BUTTONS DCGUI_BUTTON_OK      ;
		APPWINDOW MAINWINDOW():DRAWINGAREA

	IF !XSTATUS
		close all
		RETURN .T.
	ENDIF

	close all

	IF UseDb({'REBATEUP'},1,.F.)
		REBATEUP->(DBPACK())
	ELSE
	    RETURN NIL
	ENDIF



RETURN .t.

static function _REBATEADD()
	LOCAL GETLIST:={},oRebRecord, GETOPTIONS,nRec:=0  ,xStatus
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE SAYFONT '10.Courier New' GETFONT '10.Courier New'

   oRebRecord:=REBATEUP->(DC_DBRECORD(): NEW())
	REBATEUP->(DB_INIT(oRebRecord))

	@ 2,0 DCSAY 'Rebate Code to set up' get oRebRecord:rebcode  PICTURE '@!'
	@ 4,0 DCSAY 'Description          ' get oRebRecord:rebdesc
	@ 5,0 DCSAY 'Rebate Value         ' get oRebRecord:amount PICTURE '99999.99'
	dcread gui modal enterexit to xstatus options getoptions fit title 'Establish Rebate Record' eval{|o| setappwindow(o)}
	IF !xStatus
		return TRUE
	ENDIF

	IF REBATEUP->(DC_ADDREC(3,.T.,2))
		nRec:=REBATEUP->(RECNO())
		REPLACE REBATEUP->REBCODE WITH oRebRecord:rebcode
		REPLACE REBATEUP->REBDESC WITH oRebRecord:rebdesc
		REPLACE REBATEUP->AMOUNT WITH  oRebRecord:amount
		REBATEUP->(DBRUNLOCK(nRec))
	ENDIF
	REBATEUP->(DBGOTO(nRec))

RETURN NIL

static function _REBATEDEL(nRec)
	LOCAL GETLIST:={}, GETOPTIONS  ,xStatus
	DCGETOPTIONS SAYWIDTH 0 AUTORESIZE SAYFONT '10.Courier New' GETFONT '10.Courier New'
	SELECT REBATEUP
	@ 2,0 DCSAY 'Rebate Code to Delete' get  REBATEUP->rebcode  EDITPROTECT{|| TRUE }
	@ 4,0 DCSAY 'Description          ' get  REBATEUP->rebdesc  EDITPROTECT{|| TRUE }
	dcread gui modal ADDBUTTONS enterexit to xstatus options getoptions fit title 'Establish Rebate Record' eval{|o| setappwindow(o)}
	IF !xStatus
		return NIL
	ENDIF

	REBATEUP->(BV_BLANK(.F.,.T.))

RETURN NIL

static function _REBATELIST()
	LOCAL nOrient:=2, nDefmode:=4,nCopies:=1, cOutfile:='',cSentfont:='', nPage:=0 , oPrinter
	cSentfont:='10.Courier New'
	TOP_MAR(2)
	BOT_MAR(2)
	PAGE_LEN(62)


	IF !Print_Choice( 'Rebate List', @nCopies,' ',.F.,@nOrient,@cSentfont,@nDefMode,cOutfile )
			 RETURN NIL
	ENDIF

	IF nOrient=1
			PAGE_LEN(80)
	ENDIF

	IF nDefMode=8
		PAGE_LEN(32000)
	ENDIF


	IF !PrintOn('Rebate List', oPrinter, nOrient, cSentfont, nCopies)
		  RETURN NIL
	ENDIF

	SELECT REBATEUP
	REBATEUP->(DBGOTOP())

	_RebHdr(@nPage)

	DO WHILE !REBATEUP->(EOF())
		IF DCPAGEEJECT()
			_RebHdr(@nPage)
		ENDIF
	   @ DC_PRINTERROW()+1,0 DCPRINT SAY REBATEUP->RebCode
	   @ DC_PRINTERROW(),12  DCPRINT SAY REBATEUP->RebDesc
	   @ DC_PRINTERROW(),70 DCPRINT SAY  REBATEUP->Amount
		REBATEUP->(DBSKIP(1))
	ENDDO

	PRINTOFF()
	REBATEUP->(DBGOTOP())
	RETURN NIL


	static function _RebHdr(nPage)
	nPage++
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Rebate List Page ' + alltrim(str(nPage))
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Rebate Code'
	@ DC_PRINTERROW(),12  DCPRINT SAY 'Description'
	@ DC_PRINTERROW(),70 DCPRINT SAY  'Amount'
	RETURN NIL







STATIC FUNCTION ADDINT()
	LOCAL GETLIST:={}
	LOCAL GETOPTIONS , xStatus
	local cVehInterest,cFranchise
	local mcon:='Y'
	DCGETOPTIONS SAYWIDTH 0 SAYFONT '12.Courier New' getfont '12.Courier New' autoresize
	cVehInterest:=space(10)
	cFranchise:=space(2)
	@ 10,0 DCSAY 'Enter Interest Class' get cVehInterest
	@ 11,0 DCSAY 'Enter Franchise for this Class' get cFranchise
	@ 13,0 DCSAY 'OK to Add ' get mcon picture '@! L'
	dcread gui enterexit to xstatus APPWINDOW MAINWINDOW():DRAWINGAREA fit options getoptions title 'Add Interest Class'
	if !xstatus
		return nil
	endif
	if mcon='Y'
		select classort
		if classort->(dc_addrec(3,.T.,2))
			replace classort->classtype with cVehInterest
			replace classort->franchise with cFranchise
			unlock
		endif
	endif
return nil

STATIC FUNCTION delINT(rec)
	LOCAL GETLIST:={}
	LOCAL GETOPTIONS , xStatus
	local mcon:='Y'
	DCGETOPTIONS SAYWIDTH 0 SAYFONT '12.Courier New' getfont '12.Courier New' autoresize
	classort->(dbgoto(rec))
	@ 10,0 DCSAY 'Interest Class to Delete' + classort->classtype
	@ 13,0 DCSAY 'OK to Delete ' get mcon picture '@! L'
	dcread gui enterexit to xstatus APPWINDOW MAINWINDOW():DRAWINGAREA fit options getoptions title 'Delete Interest Class'
	if !xstatus
		return nil
	endif
	if mcon='Y'
		select classort
		CLASSORT->(BV_BLANK(.T.,.F.))
	endif
	CLASSORT->(DBGOTOP())
return nil


static function loadnprod(cVehInterest,cFranchise)
	cVehInterest:=classort->classtype
	cFranchise:=classort->franchise
return .T.


// END NPRODLK

FUNCTION BROWGETZIP(oRecord)
	LOCAL GETLIST:={}, aPres, oBrowse, oToolBar,cSearchKey:=space(20)
	LOCAL XSTATUS:=.T. , getoptions
	DCGETOPTIONS ;
		SAYWIDTH 0 ;
		SAYFONT '10.Courier New'  GETFONT '10.Courier New'  ;
		AUTORESIZE
	SELECT USAZIP
	ORDSETFOCUS('USAZIPNM')


	USAZIP->(DC_SETSCOPE(0,'A'))
	USAZIP->(DC_SETSCOPE(1,'Z'))
	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, -1 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 15 }  }              /* Cell Height */

@ 1,10 DCSAY 'Enter Search Key ' get cSearchkey GETID 'SEARCH' PICTURE '@!' ;
		      KEYBLOCK {|a,b,o| DBSELECTAREA('USAZIP'), DC_BrowseAutoSeek(a,o,oBrowse)}


/* ----- Create browse ----- */

	@ 3,30 DCBROWSE oBrowse ALIAS 'USAZIP'                 ;
		SIZE 80,25                                     ;
		PRESENTATION aPres  ;
		FREEZELEFT {1,2} ;
		DATALINK{|| oRecord:ZIPCODE:=USAZIP->ZIP,;
		            oRecord:CITYSTATE:=USAZIP->CITY,;
						oRecord:STATE:=USAZIP->STATE,;
						oRecord:COUNTY:=USAZIP->COUNTY,;
		            DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST)}

	DCBROWSECOL FIELD USAZIP->CITY                     ;
		HEADER "City" HCOLOR 1,BD_LINEN PARENT oBrowse                  ;
		PROTECT {|| .T.} ;
		WIDTH 30

	DCBROWSECOL FIELD USAZIP->STATE                     ;
		HEADER "State" HCOLOR 1,0 PARENT oBrowse                  ;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD USAZIP->ZIP ;
		HEADER 'ZipCode' PARENT oBrowse ;
		PROTECT {|| .T.}

	DCBROWSECOL FIELD USAZIP->COUNTY ;
		HEADER 'County' PARENT oBrowse ;
		PROTECT {|| .T.} ;
		WIDTH 15


	DCREAD GUI ;
		FIT ;
		to xStatus ;
		APPWINDOW MAINWINDOW():DRAWINGAREA   ;
		ADDBUTTONS ;
		OPTIONS GETOPTIONS  ;
		FIT ;
		TITLE 'Zip Code Browser'

  IF !xStatus
	RETURN NIL
  ENDIF

  oRecord:ZIPCODE:=USAZIP->ZIP
  oRecord:CITYSTATE:=USAZIP->CITY
  oRecord:STATE:=USAZIP->STATE

ReTURN NIL

*** END OF EXAMPLE ***



FUNCTION CHKINT(nInterestRate,nRate24,nrate36,nRate48,nRate60)
	IF nInterestRate=0
		DC_WINALERT('Zero percent Financing!!!?')
	ENDIF
	IF nRate24 < .01
		nRate24:=nInterestRate
	ENDIF
	IF nRate36 < .01
		nRate36:=nInterestRate
	ENDIF
	IF nRate48 < .01
		nRate48:=nInterestRate
	ENDIF
	IF nRate60 < .01
		nRate60:=nInterestRate
	ENDIF
  return nil




RETURN .T.

FUNCTION CHKFINAMT(P)
	IF P=0
		RETURN .F.
	ENDIF

RETURN .T.


*QUICK.PRG QUICK Payment;Calculator


FUNCTION QUICK(P,nPay24Mos,nPay36Mos,nPay48Mos,nPay60Mos,nPayodd,nTermodd,nInterestRate,nDownPay,nQuotePrice,lChng,nTaxrate,;
		nRate24,nrate36,nRate48,nRate60)
	LOCAL GETLIST:={} , cProgram , xStatus

	LOCAL GETOPTIONS,ccon,clLoadPayment:='N',nLoad:=0,pf:=0
	MEMVAR MCON,EXT,N,MOB,AP,E,V,AN,AN1,I;
		,PP,D,SPB
	default nInterestRate:=0
	default nTermodd:=0
	default nPayodd:=0
	default nDownPay:=0
	default nQuotePrice:=0
	default nTaxrate:=0
	default nRate24:=0
	default nRate36:=0
	default nRate48:=0
	default nRate60:=0



	DCGETOPTIONS AUTORESIZE SAYWIDTH 0 sayfont '10.Courier New' getfont '10.Courier New'
	cProgram:='QUICK'
	IF !MANAGER()
		DC_WINALERT('Security Violation-You Cannot Use This Function')
		RETURN NIL
	ENDIF
	/// set up

	DCGETOPTIONS AUTORESIZE SAYWIDTH 0 sayfont '10.Courier New' getfont '10.Courier New'


	@ 2,15 DCSAY 'Quick Payment Calculator'
	If lChng                                      //// JUST A CALC- LET ENTER TO P
		@ 5,5 DCSAY 'Enter Financed Amount' Get P Picture "999999" Valid{||Chkfinamt(P)}  Saycolor 9,0 Getcolor 10,0 Getfont '12.Courier'
	Else                                         /// FED FROM UPQUOTE
		@ 5,5 DCSAY 'Financed Amount   ' get P picture '999999.99' NOTABSTOP EDITPROTECT{|| .t.}
	Endif
	@ 7,5 DCSAY 'Interest Rate' Get nInterestRate Valid{||Chkint(nInterestRate,@nRate24,@nRate36,@nRate48,@nRate60)} Picture '999.999' Saycolor 9,0 Getcolor 10,0 Getfont '12.Courier'
	@ 8,5 DCSAY 'Enter Non Standard Term If Any' Get nTermodd Picture '999' Saycolor 9,0 Getcolor 10,0 Getfont '12.Courier'
	@ 9,5 DCSAY 'Load Financed by Specified Amount Y/N  ' GET clLoadPayment picture '@! L'
	@ 10,5 DCSAY 'Amount to Add to Financed Amount       ' get nLoad picture '9999'  valid{|| iif(clLoadPayment='Y',pf:=p+nload,pf:=p),;		 /// add load if there is one
	calcpayment(60,nRate60,@nPay60Mos,Pf),;
	calcpayment(48,nRate48,@nPay48Mos,Pf),;
	calcpayment(36,nRate36,@nPay36Mos,Pf),;
	calcpayment(24,nRate24,@nPay24Mos,Pf),;
	calcpayment(ntermodd,nInterestRate,@npayodd,Pf),;
	dc_getrefresh(getlist),.t.}

	@ 11,5 DCSAY 'Rate24' get nRate24 VALID {|| calcpayment(24,nRate24,@nPay24Mos,Pf),dc_getrefresh(getlist),.t.}
	@ 11,30 DCSAY 'Rate36' Get nrate36 VALID {|| calcpayment(36,nRate36,@nPay36Mos,Pf),dc_getrefresh(getlist),.t.}
	@ 11,55 DCSAY 'Rate48' get nRate48 VALID {|| calcpayment(48,nRate48,@nPay48Mos,Pf),dc_getrefresh(getlist),.t.}
	@ 11,80 DCSAY 'Rate60' get nRate60 VALID {|| calcpayment(60,nRate60,@nPay60Mos,Pf),dc_getrefresh(getlist),.t.}

	@ 13,5 DCSAY 'NS Payment Term'
	@ 13,25 DCGET nTermodd PICTURE '999'  NOTABSTOP EDITPROTECT{|| .t.}
	@ 13,35 DCGET npayodd  picture '99999.99' NOTABSTOP EDITPROTECT{|| .t.}
	@ 15,5  DCSAY '60 Months' get nPay60Mos  picture '99999.99' NOTABSTOP EDITPROTECT{|| .t.}
	@ 17,5 DCSAY '48 Months'  get nPay48Mos  picture '99999.99' NOTABSTOP EDITPROTECT{|| .t.}
	@ 19,5 DCSAY '36 Months'  get nPay36Mos  picture '99999.99' NOTABSTOP EDITPROTECT{|| .t.}
	@ 21,5 DCSAY '24 Months'  get nPay24Mos  picture '99999.99' NOTABSTOP EDITPROTECT{|| .t.}
	@ 22,20 DCSAY 'Press Alt P for Hard Copy.'
	MCON:='N'
	@ 23,20 DCSAY 'Press Cancel Button OK Button or ESC Key to Exit' ////GET MCON PICTURE "@! L"

	IF !lChng                                    /// DON'T SHOW BUTTONS IF THIS IS JUST A CALC

		@ 21,50 DCPUSHBUTTON CAPTION 'GTP 24' ;
			SIZE 10,1;
			CARGO 'CANCEL' ;
			ACTION {|| upgettopay(24,nInterestRate,@Pf,@nPay24Mos,@nDownPay,@nQuotePrice,@P,nTaxrate),;
			calcpayment(24,nRate24,@nPay24Mos,Pf),;
			calcpayment(36,nRate36,@nPay36Mos,Pf),;
			calcpayment(48,nRate48,@nPay48Mos,Pf),;
			calcpayment(60,nRate60,@nPay60Mos,Pf),;
			calcpayment(ntermodd,nInterestRate,@npayodd,Pf),;
			dc_getrefresh(getlist)}

		@ 19,50 DCPUSHBUTTON CAPTION 'GTP 36' ;
			SIZE 10,1;
			CARGO 'CANCEL' ;
			ACTION {|| upgettopay(36,nInterestRate,@Pf,nPay36Mos,@nDownPay,@nQuotePrice,@P,nTaxrate),;
			calcpayment(24,nRate24,@nPay24Mos,Pf),;
			calcpayment(36,nRate36,@nPay36Mos,Pf),;
			calcpayment(48,nRate48,@nPay48Mos,Pf),;
			calcpayment(60,nRate60,@nPay60Mos,Pf),;
			calcpayment(ntermodd,nInterestRate,@npayodd,Pf),;
			dc_getrefresh(getlist)}

		@ 17,50 DCPUSHBUTTON CAPTION 'GTP 48' ;
			SIZE 10,1;
			CARGO 'CANCEL' ;
			ACTION {|| upgettopay(48,nInterestRate,@Pf,nPay48Mos,@nDownPay,@nQuotePrice,@P,nTaxrate),;
			calcpayment(24,nRate24,@nPay24Mos,Pf),;
			calcpayment(36,nRate36,@nPay36Mos,Pf),;
			calcpayment(48,nRate48,@nPay48Mos,Pf),;
			calcpayment(60,nRate60,@nPay60Mos,Pf),;
			calcpayment(ntermodd,nInterestRate,@npayodd,Pf),;
			dc_getrefresh(getlist)}

		@ 15,50 DCPUSHBUTTON CAPTION 'GTP 60' ;
			SIZE 10,1;
			CARGO 'CANCEL' ;
			ACTION {|| upgettopay(60,nInterestRate,@Pf,nPay60Mos,@nDownPay,@nQuotePrice,@P,nTaxrate),;
			calcpayment(24,nRate24,@nPay24Mos,Pf),;
			calcpayment(36,nRate36,@nPay36Mos,Pf),;
			calcpayment(48,nRate48,@nPay48Mos,Pf),;
			calcpayment(60,nRate60,@nPay60Mos,Pf),;
			calcpayment(ntermodd,nInterestRate,@npayodd,Pf),;
			dc_getrefresh(getlist)}


		@ 13,50 DCPUSHBUTTON CAPTION 'GTP ODD' ;
			SIZE 10,1;
			CARGO 'CANCEL' ;
			ACTION {|| upgettopay(nTermodd,nInterestRate,@Pf,nPayodd,@nDownPay,@nQuotePrice,@P,nTaxrate),;
			calcpayment(24,nRate24,@nPay24Mos,Pf),;
			calcpayment(36,nRate36,@nPay36Mos,Pf),;
			calcpayment(48,nRate48,@nPay48Mos,Pf),;
			calcpayment(60,nRate60,@nPay60Mos,Pf),;
			calcpayment(ntermodd,nInterestRate,@npayodd,Pf),;
			dc_getrefresh(getlist)}

	ENDIF                                        /// lChng

	@ 7,50 DCPUSHBUTTON CAPTION 'Clear All Interest Rates' ;
			SIZE 20,1;
			CARGO 'CANCEL' ;
			ACTION {|| nRate24:=nRate36:=nRate48:=nRate60:=0,DC_GETREFRESH(GETLIST)}

	DCREAD GUI modal addbuttons FIT OPTIONS GETOPTIONS TO XSTATUS title 'Payment Calcs';
		eval{|o| setappwindow(o)}
	IF !XSTATUS
		RETURN NIL
	ENDIF
RETURN NIL

// END QUICK

static function calcpayment(n,nInterestRate,nPayment,Pf)
	local getlist:={}
	local ap,exy,i,v,an,an1,PP,D,spb,mob,ext, a


	if nInterestRate=0
		nPayment:=pf/n
		return nPayment
	endif

	ap:=n
	exy:=ext:=0
	i:=nInterestRate/1200
	v:=1/(1+i)
	an:=(1-V^AP)/i
	an1:=(1-V^N)/i
	i:=(1+(EXT/30)*i)
	PP:=D:=0
	spb:=(N-AN1)/i
	mob:=(PP/100)* 2/(N+1)
	a:=AN1*((1+i)^N) + (EXT/30) * ( 1/I)
	nPayment:=(Pf*I)/(AN-I*N*D/100-I*MOB*(SPB + (AN-AN1)*A + AN1*(EXT/30) * (1/I)))
return nPayment



/// function to 'back into  payment by inputting payment and presenting a new total amount financed to reach it '
static function upgettopay(nTerm,nRate,nFinamount,nPay,nDownPay,nQuotePrice,nTaxrate,oRecord)
	local getlist:={}
	local getoptions , xStatus
	local nFactor:=0
	local nNewpay:=0
	local nNewamtfin:=0
	local nAmtreduce:=0
	local nApptosale:=0,nApptodown:=0            /// INITIALIZE VARS TO PUT nAmtreduce TO
	DCGEToptions saywidth 0 sayfont '12.Courier New' getfont '12.Courier New' autoresize

	/// get current factor
	nFactor:=npay/(nFinamount)

	@ 9,0 DCSAY  '                          Current payment ' get nPay picture '99999.99'  NOTABSTOP EDITPROTECT{|| .t.}
	@ 10,0 DCSAY '                   Current amount financed' get nfinamount picture '999999.99' NOTABSTOP EDITPROTECT{|| .t.}
	@ 11,0 DCSAY '                 Current Factor per $1000 ' get nFactor picture '.99999' NOTABSTOP EDITPROTECT{|| .t.}
	@ 12,0 DCSAY '                     Enter desired payment' get nNewpay picture '99999.99' valid{|| makenewpay(nFinamount,nNewpay,nFactor,@nNewamtfin,@nAmtreduce,nTaxrate),dc_getrefresh(getlist),.t.}
	@ 14,0 DCSAY '                     New Amount to finance' get nNewamtfin picture '999999.99' NOTABSTOP EDITPROTECT{|| .t.}
	@ 15,0 DCSAY 'Amount of Reduction needed to get payment.' get nAmtreduce picture '99999.99' NOTABSTOP EDITPROTECT{|| .t.}
	@ 18,0 DCSAY '                 Increase Down Payment by ' get nApptodown picture '99999.99'  valid{|| nApptosale:=nAmtreduce-nApptodown,dc_getrefresh(getlist),.t.}
	@ 19,0 DCSAY '              Reduce Vehicle Sale Price by' get nApptosale picture '99999.99'

	@ 22,20 dcpushbutton CAPTION 'Apply reductions to deal calculation.' ;
		SIZE 40,2;
		CARGO 'CANCEL' ;
		ACTION {|| applytodeal(nApptosale,nApptodown,@nDownPay,@nQuotePrice,@nFinamount,@oRecord),dc_getrefresh(getlist),DC_READGUIEVENT(DCGUI_EXIT_OK,GETLIST) }

	dcread gui to xstatus fit addbuttons options getoptions title 'Get to Payment' modal eval{|o| setappwindow(o)}
	if !xstatus
		return nil
	endif
return nil



/// apply reductions to sale and downpayment
static function applytodeal(nApptosale,nApptodown,nDownPay,nQuotePrice,nFinamount,oRecord)
	oRecord:Quoted:=oRecord:Quoted-nApptosale
	oRecord:DownPay:=oRecord:DOWNPAY+nApptodown
	oRecord:Finance:=oRecord:finance-(nApptosale+nApptodown) ///reset loaded fin amt
	//P:=P-(nApptosale+nApptodown)                 /// reset unloaded fin amt
return nil




/// make new net fin amount based on payment desired and factor derived from current payment term and rate.

static function makenewpay(nFinamount,nNewpay,nFactor,nNewamtfin,nAmtreduce,nTaxrate)
	nNewamtfin:=nNewpay/nFactor
	nAmtreduce:=nFinamount-nNewamtfin
	nAmtreduce:=(nAmtreduce-(nAmtreduce*(nTaxrate/100))) /// only reduce by taking out effective sales tax
return .t.

FUNCTION UPDTACES()
LOCAL GETLIST:={}, oDialog , aStruct:={} , oRecord
LOCAL lGoahead:=.f.
IF FEXISTS('ACEVEHIC2.DBF')
	ERASE ACEVEHIC2.DBF
ENDIF
IF !USE_UDF('ACEVEHIC',.T.)
	DC_WINALERT('Try Later file in use.')
else
	ACEVEHIC->(DBGOTOP())
ENDIF



oDialog := DC_WAITON('Now Converting Files')

SELECT ACEVEHIC
aStruct:=ACEVEHIC->(DBSTRUCT())


IF !ACEVEHIC->(ISFIELDVAR('REBCODE1'))
	aAdd(aStruct,{'REBCODE1','C',8,0})
	lGoahead:=.t.
ENDIF
IF !ACEVEHIC->(ISFIELDVAR('REBCODE2'))
	aAdd(aStruct,{'REBCODE2','C',8,0})
	lGoahead:=.t.

ENDIF
IF !ACEVEHIC->(ISFIELDVAR('REBCODE3'))
	aAdd(aStruct,{'REBCODE3','C',8,0})
	lGoahead:=.t.

ENDIF
IF !ACEVEHIC->(ISFIELDVAR('REBCODE4'))
	aAdd(aStruct,{'REBCODE4','C',8,0})
	lGoahead:=.t.

ENDIF
IF !ACEVEHIC->(ISFIELDVAR('REBCODE5'))
	aAdd(aStruct,{'REBCODE5','C',8,0})
	lGoahead:=.t.

ENDIF
IF !ACEVEHIC->(ISFIELDVAR('REBCODE6'))
	aAdd(aStruct,{'REBCODE6','C',8,0})
	lGoahead:=.t.

ENDIF
IF lGoahead:=.t.

 DBCREATE('ACEVEHIC2',aStruct, 'FOXCDX')
 USE ACEVEHIC2 NEW
 SELECT ACEVEHIC
 ORDSETFOCUS(0)
 ACEVEHIC->(DBGOTOP())
 DO WHILE !ACEVEHIC->(EOF())
	oRecord:=DC_DBRECORD():NEW()
	ACEVEHIC->(DC_DBSCATTER(oRecord))
	ACEVEHIC2->(DC_DBGATHER(oRecord,.t.))
	SELECT ACEVEHIC
	ACEVEHIC->(DBSKIP(1))
 ENDDO
 ACEVEHIC->(DBCLOSEAREA())
 ACEVEHIC2->(DBCLOSEAREA())
 SLEEP(100)
 COPY FILE (GETFLG('ACCTPATH')+'ACEVEHIC.DBF') TO (GETFLG('ACCTPATH')+'ACEVEHICOLD.DBF')
 COPY FILE (GETFLG('ACCTPATH')+'ACEVEHIC2.DBF') TO (GETFLG('ACCTPATH')+'ACEVEHIC.DBF')

 DC_IMPL(oDialog)
endif


 RETURN NIL


*UPDELETE.PRG PROGRAM TO DELETE UPS IN HISTORY OR CUR MONTH
FUNCTION UPDELETE()
	LOCAL GETLIST:={} , cProgram , cUpId , nRec ,clDelete:='Y'  , mcon ,  getoptions , xstatus , cLastName:=space(10) ,aStruct:={},oRecord
	cProgram:='UPDELETE'
	IF.NOT.SECURITY(@cProgram)
	 DC_winALERT('SECURITY VIOLATION')
    RETURN NIL
   ENDIF

cUPID:=SPACE(7)
cProgram:="UPDELETE"
nREC:=0
MCON := "N"

IF !UseDb({'NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF



DCGETOPTIONS TABSTOP AUTORESIZE SAYWIDTH 0  SAYFONT '10.COURIER NEW' GETFONT '10.COURIER NEW'
@ 5,0 DCSAY 'Enter a Y to Activate Up Delete Function or N to Just Change Up Data '  get clDelete picture '@! L'
@ 10,0 DCSAY 'Enter Name To Search For' GET cLastName picture '@!'
DCREAD GUI MODAL ENTEREXIT TO XSTATUS FIT OPTIONS GETOPTIONS TITLE 'ACTIVATE DELETE UP RECORD' EVAL{|O| SETAPPWINDOW(O)}

IF !XSTATUS
  RETURN NIL
ENDIF

SELECT NEWUPS
ORDSETFOCUS('NEWUPNM')
NEWUPS->(DBGOTOP())
SET SOFTSEEK ON
SEEK UPPER(cLastName)
SET SOFTSEEK OFF

/// browse the file to allow changes
SELECT NEWUPS
BROWUPS(@nREC)
/// if activate delete was selected
IF clDelete='Y'
 IF nREC=0
	DC_winALERT("NO RECORD CHOSEN")
   RETURN NIL
 ENDIF
	SELECT NEWUPS
	GOTO nREC
	@ 9,0 DCSAY 'UPID ' + NEWUPS->UPID
	@ 10,0 DCSAY NEWUPS->FIRSTNAME+' '+NEWUPS->LASTNAME
	@ 11,0 DCSAY NEWUPS->STREET
	@ 12,0 DCSAY NEWUPS->CITYSTATE
	@ 13,0 DCSAY 'DATE IN '+ DTOC(NEWUPS->DATEIN)
	@ 14,0 DCSAY NEWUPS->SALESMAN
	@ 15,0 DCSAY 'IS THIS THE CORRECT UP TO DELETE Y/N' GET clDelete PICTURE '@! L'
	DCREAD GUI ENTEREXIT TO XSTATUS FIT OPTIONS GETOPTIONS TITLE 'DELETE THIS UP RECORD'
	IF !XSTATUS
	 RETURN NIL
	ENDIF
	IF clDelete = "N"
		return nil
	ENDIF
	IF clDelete = "Y"
		nRec:= NEWUPS->(RECNO())
		SELECT NEWUPS
		GOTO nREC
		NEWUPS->(BV_BLANK())
	ENDIF
 endif /// clDelete
CLOSE DATABASES
RETURN NIL
//* END UPDELETE.PRG

*UP2UPSHT.PRG

FUNCTION UPHDR(MESSAGE,nPage)
nPage++
@ DC_PRINTERROW()+1,0 DCPRINT SAY MESSAGE FONT '10.Courier New Bold'
@ DC_PRINTERROW(),80 DCPRINT SAY 'Page '+ alltrim(str(nPage))

@ DC_PRINTERROW()+1,0 DCPRINT SAY 'SM' FONT '8.Courier New Bold'
@ DC_PRINTERROW(),5 DCPRINT SAY 'Stat'        FONT '8.Courier New Bold'
@ DC_PRINTERROW(),10 DCPRINT SAY 'Last Name'  FONT '8.Courier New Bold'
@ DC_PRINTERROW(),50 DCPRINT SAY 'First Name' FONT '8.Courier New Bold'
@ DC_PRINTERROW(),80 DCPRINT SAY 'A/C'        FONT '8.Courier New Bold'
@ DC_PRINTERROW(),85 DCPRINT SAY 'Phone'      FONT '8.Courier New Bold'
@ DC_PRINTERROW(),95 DCPRINT SAY 'WorkPhn'    FONT '8.Courier New Bold'
@ DC_PRINTERROW(),115 DCPRINT SAY 'CellPhn'   FONT '8.Courier New Bold'
//@ DC_PRINTERROW(),126 DCPRINT SAY 'Comments'  FONT '8.Courier New Bold'
@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Date In'   FONT '8.Courier New Bold'
@ DC_PRINTERROW(),11 DCPRINT SAY  'Street'    FONT '8.Courier New Bold'
@ DC_PRINTERROW(),32 DCPRINT SAY 'UpType'     FONT '8.Courier New Bold'
@ DC_PRINTERROW(),40 DCPRINT SAY 'Interest'   FONT '8.Courier New Bold'
@ DC_PRINTERROW(),50 DCPRINT SAY 'ToMGR'      FONT '8.Courier New Bold'
@ DC_PRINTERROW(),60 DCPRINT SAY 'AdSrce'   FONT '8.Courier New Bold'
@ DC_PRINTERROW(),68 DCPRINT SAY 'New/Used'   FONT '8.Courier New Bold'
@ DC_PRINTERROW(),80 DCPRINT SAY 'OrigSrce'   FONT '8.Courier New Bold'
@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Follow Up' FONT '8.Courier New Bold'
@ DC_PRINTERROW(),11 DCPRINT SAY 'CityState'  FONT '8.Courier New Bold'
@ DC_PRINTERROW(),40 DCPRINT SAY 'Zipcode'    FONT '8.Courier New Bold'
@ DC_PRINTERROW(),50 DCPRINT SAY 'eMailAddress' FONT '8.Courier New Bold'
@ DC_PRINTERROW(),101 DCPRINT SAY 'Beback1'    FONT '8.Courier New Bold'
@ DC_PRINTERROW(),110 DCPRINT SAY 'Beback2'   FONT '8.Courier New Bold'
@ DC_PRINTERROW(),125 DCPRINT SAY 'Franch' FONT '8.Courier New Bold'
@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Comments/Sold Info' FONT '8.Courier New Bold'
RETURN NIL



* BATUPSHT.PRG BATCH PROCESS UPSHEETS
FUNCTION BATUPSHT
LOCAL GETLIST:={} , getoptions , xstatus, lPrevue
LOCAL dBEGDATE,dENDDATE , cProgram
local lMakefile:=.t. , cSM , cMessage:='' , nPage:=0
local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'

cProgram:='BATUPSHT'
IF.NOT.SECURITY(@cProgram)
    DC_winALERT('SECURITY VIOLATION')
    RETURN NIL
ENDIF


dBEGDATE:=dENDDATE:=CTOD("  /  /  ")

IF !UseDb({'SM','NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF


XSTATUS:=.T.
DCGETOPTIONS AUTORESIZE SAYWIDTH 0
@ 5,5 DCSAY 'Enter Date From' GET dBEGDATE VALID{||CHKBEGDT(dBEGDATE)}
@ 7,5 DCSAY 'Enter Date To  ' GET dENDDATE
@ 12,5 DCCHECKBOX lMakefile PROMPT 'Check Here to Create Mail Merge File (upmail.dbf)'


DCREAD GUI appwindow mainwindow():drawingarea TO XSTATUS ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'BATCH UPSHEETS'
IF !XSTATUS
  RETURN NIL
ENDIF
if lmakefile             //// open upmail if necessaRY
  if use_udf('upmail',.t.)
    zap
  else
    close all
  endif
endif

IF !Print_Choice('Batch Ups', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Batch Ups ', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF



SELECT SM
DO WHILE .NOT. EOF()
  cSM:=SM->SALESMAN
  IF SM->ACTIVE='Y'
   BUPSHEET(cSM,dBEGDATE,dENDDATE,lMakefile)
   dcprint eject
	SELECT SM
   SEEK cSM
  ENDIF
	SKIP
ENDDO
printoff(oPrinter)
IF lMakefile
	DC_WINALERT('File UpMail.dbf has been created.')
ENDIF
CLOSE DATABASES
RETURN NIL

// END BATUPSHT

FUNCTION IBATUPSHT
LOCAL GETLIST:={} , getoptions
LOCAL dBEGDATE,dENDDATE,cVehInterest ,cProgram ,  xStatus, cFranchise:=''
local lMakefile:=.t. , cSM , cMessage:='' , nPage:=0
local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE
cProgram:='BATUPSHT'
IF.NOT.SECURITY(@cProgram)
    DC_WINALERT('SECURITY VIOLATION')
    RETURN NIL
  ENDIF


dBEGDATE:=dENDDATE:=CTOD("  /  /  ")

IF !UseDb({'CLASSORT','SM','NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF


XSTATUS:=.T.
cVehInterest:=space(10)
@ 5,5 DCSAY 'Enter Vehicle Interest' GET cVehInterest ;
SAYCOLOR 3,0;
VALID{||CHKINTR(@cVehInterest,@cFranchise),DC_GETREFRESH(GETLIST),.T.}
@ 7,5 DCSAY 'Enter Date From' GET dBEGDATE VALID{||CHKBEGDT(dBEGDATE)}
@ 9,5 DCSAY 'Enter Date To  ' GET dENDDATE
@ 12,5 DCCHECKBOX lMakefile PROMPT 'Check Here to Create Mail Merge File (upmail.dbf)'


DCREAD GUI appwindow mainwindow():drawingarea TO XSTATUS ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'BATCH UPSHEETS'
IF !XSTATUS
  RETURN NIL
ENDIF
if lmakefile             //// open upmail if necessaRY
  if use_udf('upmail',.t.)
    zap
  else
    close all
  endif
endif

IF !Print_Choice('Batch Ups', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Batch Ups ', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF


SELECT SM
DO WHILE .NOT. EOF()
	IF SM->ACTIVE='Y'
     cSM:=SM->SALESMAN
     INBUPSHEET(cSM,dBEGDATE,dENDDATE,lmakefile,cVehInterest)
     dcprint eject
	  SELECT SM
     SEEK cSM
	ENDIF
	SKIP
ENDDO
printoff(oPrinter)
CLOSE DATABASES
RETURN NIL

// END BATUPSHT

FUNCTION fBATUPSHT
LOCAL GETLIST:={} , cSM  , xStatus , lPrevue
LOCAL dBEGDATE,dENDDATE,cFranchise , getoptions ,cProgram
local lMakefile:=.t. ,  cMessage:='' , nPage:=0
local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
cProgram:='BATUPSHT'
IF.NOT.SECURITY(@cProgram)
    DC_WINALERT('SECURITY VIOLATION')
    RETURN NIL
ENDIF

dBEGDATE:=dENDDATE:=CTOD("  /  /  ")

IF !UseDb({'SM','NEWUPS',''})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF

cFranchise:=space(10)
XSTATUS:=.T.
DCGETOPTIONS AUTORESIZE SAYWIDTH 0
@ 5,5 DCSAY 'Enter Franchise' GET cFranchise ;
SAYCOLOR 3,0
@ 7,5 DCSAY 'Enter Date From' GET dBEGDATE VALID{||CHKBEGDT(dBEGDATE)}
@ 9,5 DCSAY 'Enter Date To  ' GET dENDDATE
@ 12,5 DCCHECKBOX lMakefile PROMPT 'Check Here to Create Mail Merge File (upmail.dbf)'


DCREAD GUI appwindow mainwindow():drawingarea TO XSTATUS ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'BATCH UPSHEETS'
IF !XSTATUS
  CLOSE DATABASES
  RETURN NIL
ENDIF
if lmakefile             //// open upmail if necessaRY
  if use_udf('upmail',.t.)
    zap
  else
    close all
  endif
endif

IF !Print_Choice('Batch Ups', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Batch Ups ', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF


SELECT SM
DO WHILE .NOT. EOF()
	IF SM->ACTIVE='Y'
    cSM:=SM->SALESMAN
    fBUPSHEET(cSM,dBEGDATE,dENDDATE,lmakefile,cFranchise)
    dcprint eject
	 SELECT SM
    SEEK cSM
	ENDIF
	SKIP
ENDDO
printoff(oPrinter)
CLOSE DATABASES
RETURN NIL

// END BATUPSHT


FUNCTION INDVUPSHT
LOCAL GETLIST:={}
LOCAL dBEGDATE,dENDDATE,cSM , cProgram , xStatus
LOCAL GETOPTIONS, nPage:=0
local lmakefile:=.f. , cMessage:=''
local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'

BOT_MAR(2)
TOP_MAR(2)
PAGE_LEN(66)
DCGETOPTIONS AUTORESIZE SAYWIDTH 0 SAYFONT '12.Courier New' getfont '12.Courier New'
cProgram:='BATUPSHT'
IF.NOT.SECURITY(@cProgram)
    DC_winALERT('SECURITY VIOLATION')
    RETURN NIL
  ENDIF
dBEGDATE:=dENDDATE:=CTOD("  /  /  ")

cSM:=space(3)

IF !UseDb({'CLASSORT','SM','NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF

XSTATUS:=.T.
@ 10,5 DCSAY 'Enter Sales Person to Print For'  get cSM picture '@!'    valid {|| chksm(cSM)}
@ 12,5 DCSAY 'Enter Date From' GET dBEGDATE VALID{||CHKBEGDT(dBEGDATE)}
@ 14,5 DCSAY 'Enter Date To  ' GET dENDDATE
@ 16,5 DCCHECKBOX lMakefile PROMPT 'Check Here to Create Mail Merge File (upmail.dbf)'
DCREAD GUI appwindow mainwindow():drawingarea TO XSTATUS addbuttons  ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'INDIVIDUAL UPSHEETS'
IF !XSTATUS
  RETURN NIL
ENDIF
if lmakefile             //// open upmail if necessaRY
  if use_udf('upmail',.t.)
    zap
  else
    close all
  endif
endif

IF !Print_Choice('Ups By SM', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Ups By SM', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF

cMessage:='INDIVIDUAL UPSHEET '+ DTOC(dBEGDATE)+' TROUGH ' + DTOC(dENDDATE)
UPHDR(cMessage,@nPage)
DO WHILE !NEWUPS->(EOF())
  if NEWUPS->salesman=cSM
      if NEWUPS->datein >= dbegdate
        if NEWUPS->datein <= denddate
          if lmakefile                                        /// make file if Y  upmail.dbf
            makeupmail()
          endif

			 _PrintUprecord(cMessage,@nPage)


        endif ///begdate
      endif  //enddate
    endif /// smi
  SKIP ALIAS NEWUPS
ENDDO
PRINTOFF(oPrinter)
close all
if lmakefile
	DC_WINALERT('Data File for eMail has been created.  upmail.dbf')
endif
RETURN NIL

// END INDVUPSHT

FUNCTION INDVINTUPS
LOCAL GETLIST:={}
LOCAL dBEGDATE,dENDDATE,cSM,cVehInterest , cFranchise , xStatus
LOCAL GETOPTIONS , cProgram, nPage:=0
local lmakefile:=.f. , cMessage:=''
local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'


BOT_MAR(2)
TOP_MAR(2)
PAGE_LEN(66)
DCGETOPTIONS AUTORESIZE SAYWIDTH 0 SAYFONT '12.Courier New' getfont '12.Courier New'
cProgram:='BATUPSHT'
IF.NOT.SECURITY(@cProgram)
    DC_winALERT('SECURITY VIOLATION')
    CLOSE DATABASES
    RETURN NIL
  ENDIF
dBEGDATE:=dENDDATE:=CTOD("  /  /  ")
cSM:=space(2)
cVehInterest:=space(10)

IF !UseDb({'CLASSORT','SM','NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF


cFranchise:=' '

XSTATUS:=.T.
DCGETOPTIONS AUTORESIZE SAYWIDTH 0
@ 2,5 DCSAY 'ENTER INTEREST' GET cVehInterest ;
SAYCOLOR 3,0;
VALID{||CHKINTR(@cVehInterest,@cFranchise),DC_GETREFRESH(GETLIST),.T.}
@ 5,5   DCSAY 'Enter Sales Person to Print For'  get cSM picture '@!'    valid {|| chksm(cSM)}
@ 7,5 DCSAY 'Enter Date From' GET dBEGDATE VALID{||CHKBEGDT(dBEGDATE)}
@ 9,5 DCSAY 'Enter Date To  ' GET dENDDATE
@ 16,5 DCCHECKBOX lMakefile PROMPT 'Check Here to create Mail Merge File (upmail.dbf)'

DCREAD GUI appwindow mainwindow():drawingarea TO XSTATUS addbuttons  ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'INDIVIDUAL UPSHEETS'
IF !XSTATUS
  CLOSE DATABASES
  RETURN NIL
ENDIF
if lmakefile             //// open upmail if necessaRY
  if use_udf('upmail',.t.)
    zap
  else
    close all
  endif
endif

IF !Print_Choice('Ups By SM', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Ups By SM', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF

cMessage:='INDIVIDUAL UPSHEET Interest '+cVehInterest+'  '+cSM+'  '+ DTOC(dBEGDATE)+' TROUGH ' + DTOC(dENDDATE)
UPHDR(cMessage,@nPage)
DO WHILE !NEWUPS->(EOF())
  if NEWUPS->salesman=cSM
      if NEWUPS->datein >= dbegdate
        if NEWUPS->datein <= denddate
          IF NEWUPS->INTEREST=cVehInterest
            if lmakefile                                        /// make file if Y  upmail.dbf
              makeupmail()
            endif

				_PrintUprecord(cMessage,@nPage)

          ENDIF   /// cVehInterest
        endif ///begdate
      endif  //enddate
    endif /// smi
  SKIP ALIAS NEWUPS
ENDDO
PRINTOFF(oPrinter)
close all
if lmakefile
	DC_WINALERT('Data File for Mail Merge has been created.  upmail.dbf')
endif

RETURN NIL

// END INDVUPSHT

static function makeupmail()           /// add record to mail file
IF NEWUPS->SALECODE=0             /// MAKE SURE NONE ARE SOLD
  select upmail
  if upmail->(dc_addrec())
    replace upmail->lastname with NEWUPS->lastname
    replace upmail->firstname with NEWUPS->firstname
    replace upmail->street with NEWUPS->street
    replace upmail->citystate with NEWUPS->citystate
    replace upmail->zipcode with NEWUPS->zipcode
    replace upmail->interest with NEWUPS->interest
    replace upmail->datein with NEWUPS->datein
    replace upmail->salesman with NEWUPS->salesman
    replace upmail->email with NEWUPS->email
    UPMAIL->(dbcommit())
  endif
ENDIF
return nil


FUNCTION INTRUPSHT
LOCAL GETLIST:={}   , cProgram  , cFranchise , xStatus , nPage:=0
LOCAL dBEGDATE,dENDDATE,cVehInterest
LOCAL GETOPTIONS
local lmakefile:=.f.  , cMessage:=''
local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
BOT_MAR(2)
TOP_MAR(2)
PAGE_LEN(66)
DCGETOPTIONS AUTORESIZE SAYWIDTH 0 SAYFONT '12.Courier New' getfont '12.Courier New'
cProgram:='BATUPSHT'
IF.NOT.SECURITY(@cProgram)
    DC_winALERT('SECURITY VIOLATION')
    RETURN NIL
  ENDIF
dBEGDATE:=dENDDATE:=CTOD("  /  /  ")
cVehInterest:=space(10)
cFranchise:=' '

IF !UseDb({'CLASSORT','SM','NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF



XSTATUS:=.T.
DCGETOPTIONS AUTORESIZE SAYWIDTH 0
@ 2,5 DCSAY 'Vehicle Interest' GET cVehInterest ;
SAYCOLOR 3,0;
VALID{||CHKINTR(@cVehInterest,@cFranchise),DC_GETREFRESH(GETLIST),.T.}
@ 5,5 DCSAY 'Enter Date From' GET dBEGDATE VALID{||CHKBEGDT(dBEGDATE)}
@ 7,5 DCSAY 'Enter Date To  ' GET dENDDATE
@ 9,5 DCCHECKBOX lMakefile PROMPT 'Check Here to create Mail Merge File (upmail.dbf)'

DCREAD GUI appwindow mainwindow():drawingarea TO XSTATUS addbuttons  ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Interest Upsheets'
IF !XSTATUS
  CLOSE DATABASES
  RETURN NIL
ENDIF
if lmakefile             //// open upmail if necessaRY
  if use_udf('upmail',.t.)
    zap
  else
    close all
  endif
endif

IF !Print_Choice('Ups By SM', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Ups By SM', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF



cMessage:='INDIVIDUAL UPSHEET Interest '+cVehInterest+' '+ DTOC(dBEGDATE)+' TROUGH ' + DTOC(dENDDATE)
UPHDR(cMessage,@nPage)
DO WHILE !NEWUPS->(EOF())
  if NEWUPS->INTEREST=cVehInterest
      if NEWUPS->datein >= dbegdate
        if NEWUPS->datein <= denddate
            if lmakefile                                        /// make file if Y  upmail.dbf
             makeupmail()
            endif


				_PrintUprecord(cMessage,@nPage)

        endif ///begdate
      endif  //enddate
    endif /// smi
  SKIP ALIAS NEWUPS
ENDDO
DCPRINT OFF
close all
IF lMakefile
	DC_WINALERT('File UpMail.dbf has been Created.')
ENDIF

RETURN NIL


FUNCTION StatUpRept()
LOCAL GETLIST:={}   , cProgram  , cFranchise , xStatus , cStatus , cSM:=space(3)
LOCAL dBEGDATE,dENDDATE, nPage:=0  , cMessage:='' , oStatButtons
LOCAL GETOPTIONS
local lmakefile:=.f.
local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
BOT_MAR(2)
TOP_MAR(2)
PAGE_LEN(66)
DCGETOPTIONS AUTORESIZE SAYWIDTH 0 SAYFONT '12.Courier New' getfont '12.Courier New'
cProgram:='BATUPSHT'
IF.NOT.SECURITY(@cProgram)
    DC_winALERT('SECURITY VIOLATION')
    RETURN NIL
  ENDIF
dBEGDATE:=dENDDATE:=CTOD("  /  /  ")
cStatus:='A'
cFranchise:=' '

IF !UseDb({'NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF


XSTATUS:=.T.
@ 1,5 DCGROUP oStatButtons CAPTION 'Select a Status to Print'  SIZE 35,12
	  @ 1,2 DCRADIO cStatus VALUE 'A' PROMPT '     Active      ' COLOR 1,0 ;
      PARENT oStatButtons
	  @ 3,2 DCRADIO cStatus VALUE 'F' PROMPT 'Future Follow up ' COLOR 1,15 ;
      PARENT oStatButtons
	  @ 5,2 DCRADIO cStatus VALUE 'R' PROMPT '     Rescue      ' COLOR 2,3 ;
      PARENT oStatButtons
	  @ 7,2 DCRADIO cStatus VALUE 'S' PROMPT '     Sold        ' COLOR 1,5 ;
      PARENT oStatButtons
	  @ 9,2 DCRADIO cStatus VALUE 'D' PROMPT '     Dead        ' COLOR 0,1 ;
      PARENT oStatButtons


@ 14,5 DCSAY 'Enter Date From' GET dBEGDATE VALID{||CHKBEGDT(dBEGDATE)}
@ 15,5 DCSAY 'Enter Date To  ' GET dENDDATE
@ 16,0 DCSAY 'Enter a Sales Person Initials or leave blank for all.' get cSM picture '@!'
@ 17,5 DCCHECKBOX lMakefile PROMPT 'Check Here to Create Mail Merge File (upmail.dbf)'
DCREAD GUI appwindow mainwindow():drawingarea TO XSTATUS ADDbuttons ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Interest Upsheets'
IF !XSTATUS
  CLOSE DATABASES
  RETURN NIL
ENDIF
if lmakefile             //// open upmail if necessaRY
  if use_udf('upmail',.t.)
    zap
  else
    close all
  endif
endif

IF !Print_Choice('Ups By Status', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Ups By Status', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF

cMessage:='UpStatus '+cStatus+'  '+ DTOC(dBEGDATE)+' TROUGH ' + DTOC(dENDDATE)
UPHDR(cMessage,@nPage)

IF !empty(cSM)
	NEWUPS->(DBSETFILTER({|| NEWUPS->SALESMAN=cSM}))
ENDIF

DO WHILE !NEWUPS->(EOF())
  if NEWUPS->UPSTATUS=cStatus
      if NEWUPS->datein >= dbegdate
        if NEWUPS->datein <= denddate
            if lmakefile                                        /// make file if Y  upmail.dbf
             makeupmail()
            endif

				_PrintUprecord(cMessage,@nPage)


        endif ///begdate
      endif  //enddate
    endif /// status
  SKIP ALIAS NEWUPS
ENDDO
PRINTOFF(oPrinter)

IF lMakefile
	DC_WINALERT('File UpMail.dbf has been Created.')
ENDIF

close all
RETURN NIL

static function _PrintUprecord(cMessage,nPage)
	@ DC_PRINTERROW()+1,0 DCPRINT SAY NEWUPS->SALESMAN  FONT '8.COURIER BOLD'
	@ DC_PRINTERROW(),5 DCPRINT SAY NEWUPS->Upstatus
   @ DC_PRINTERROW(),10 DCPRINT SAY NEWUPS->LASTNAME   FONT '8.COURIER BOLD'
   @ DC_PRINTERROW(),51 DCPRINT SAY NEWUPS->FIRSTNAME
   @ DC_PRINTERROW(),80 DCPRINT SAY NEWUPS->AREA
   @ DC_PRINTERROW(),85 DCPRINT SAY NEWUPS->UPID
   @ DC_PRINTERROW(),95 DCPRINT SAY NEWUPS->WORKPHN
   @ DC_PRINTERROW(),115 DCPRINT SAY NEWUPS->CELLPHONE
   IF DCPAGEEJECT()
       UPHDR(cMessage,@nPage)
   ENDIF
	@ DC_PRINTERROW()+1,0 DCPRINT SAY '  '
   @ DC_PRINTERROW(),0 DCPRINT SAY NEWUPS->DATEIN FONT '8.COURIER BOLD'
   @ DC_PRINTERROW(),11 DCPRINT SAY NEWUPS->STREET
   @ DC_PRINTERROW(),36 DCPRINT SAY NEWUPS->UPTYPE
   @ DC_PRINTERROW(),38 DCPRINT SAY NEWUPS->INTEREST
   @ DC_PRINTERROW(),50 DCPRINT SAY NEWUPS->TOMGR
   @ DC_PRINTERROW(),60 DCPRINT SAY NEWUPS->ADSOURCE
   @ DC_PRINTERROW(),70 DCPRINT SAY NEWUPS->NEWUSED
	@ DC_PRINTERROW(),85 DCPRINT SAY NEWUPS->ORIGSOURCE FONT '8.COURIER BOLD'

	IF DCPAGEEJECT()
       UPHDR(cMessage,@nPage)
   ENDIF
	@ DC_PRINTERROW()+1,0 DCPRINT SAY '  '
   @ DC_PRINTERROW(),0 DCPRINT SAY NEWUPS->TICKLE
   @ DC_PRINTERROW(),11 DCPRINT SAY NEWUPS->CITYSTATE
   @ DC_PRINTERROW(),40 DCPRINT SAY NEWUPS->ZIPCODE
   @ DC_PRINTERROW(),50 DCPRINT SAY substr(NEWUPS->EMAIL,1,50)
   @ DC_PRINTERROW(),101 DCPRINT SAY NEWUPS->BEBACK1
   @ DC_PRINTERROW(),110 DCPRINT SAY NEWUPS->BEBACK2
   @ DC_PRINTERROW(),125 DCPRINT SAY NEWUPS->FRANCHISE FONT '8.COURIER BOLD'
	IF DCPAGEEJECT()
       UPHDR(cMessage,@nPage)
   ENDIF
	/// PRINT COMMENTS IF ANY
	IF !EMPTY(NEWUPS->COMMENTS)
	 @ DC_PRINTERROW()+1,0 DCPRINT SAY NEWUPS->COMMENTS
	 IF DCPAGEEJECT()
       UPHDR(cMessage,@nPage)
    ENDIF
	ENDIF
	IF !EMPTY(NEWUPS->COMMENT2)
	 @ DC_PRINTERROW()+1,0 DCPRINT SAY NEWUPS->COMMENT2
	 IF DCPAGEEJECT()
       UPHDR(cMessage,@nPage)
    ENDIF
	ENDIF

   IF NEWUPS->SALECODE > 0
     @ DC_PRINTERROW()+1,0 DCPRINT SAY '  '
     @ DC_PRINTERROW(),0 DCPRINT SAY 'SOLD!!' FONT '10.COURIER BOLD'
     @ DC_PRINTERROW(),10 DCPRINT SAY NEWUPS->SALECODE  FONT '8.COURIER BOLD'
     @ DC_PRINTERROW(),20 DCPRINT SAY NEWUPS->STOCKNUM  FONT '8.COURIER BOLD'
     @ DC_PRINTERROW(),30 DCPRINT SAY NEWUPS->VEHQUOTED  FONT '8.COURIER BOLD'
	  IF DCPAGEEJECT()
       UPHDR(cMessage,@nPage)
     ENDIF
   ENDIF
	skipaline()
	IF DCPAGEEJECT()
       UPHDR(cMessage,@nPage)
   ENDIF

	return nil



// END INDVUPSHT
FUNCTION iNetups
LOCAL GETLIST:={}, cProgram , cMess , xStatus , oSrcButtons , cSrc
LOCAL dBEGDATE,dENDDATE , cMessage:='' , nPage:=0
local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
LOCAL GETOPTIONS
BOT_MAR(2)
TOP_MAR(2)
PAGE_LEN(66)
DCGETOPTIONS AUTORESIZE SAYWIDTH 0 SAYFONT '12.Courier New' getfont '12.Courier New'
cProgram:='BATUPSHT'
IF.NOT.SECURITY(@cProgram)
    DC_winALERT('SECURITY VIOLATION')
    RETURN NIL
  ENDIF
dBEGDATE:=dENDDATE:=CTOD("  /  /  ")

IF !UseDb({'CLASSORT','NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF



XSTATUS:=.T.

@ 5,5 DCGROUP oSrcButtons CAPTION 'Select a Source to Print'  SIZE 35,8
	  @ 1,2 DCRADIO cSrc VALUE 'I' PROMPT '     Internet      ' COLOR 1,0 ;
      PARENT oSrcButtons
	  @ 3,2 DCRADIO cSrc VALUE 'F' PROMPT '    Floor Traffic  ' COLOR 1,15 ;
      PARENT oSrcButtons
	  @ 5,2 DCRADIO cSrc VALUE 'P' PROMPT '     Phone Up      ' COLOR 2,3 ;
      PARENT oSrcButtons
@ 12,23 DCSAY 'Enter Date From ' GET dBEGDATE VALID{||CHKBEGDT(dBEGDATE)}
@ 14,23 DCSAY 'Enter Date To   ' GET dENDDATE
DCREAD GUI appwindow mainwindow():drawingarea TO XSTATUS buttons 2 ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Interest Upsheets'
IF !XSTATUS
  CLOSE DATABASES
  RETURN NIL
ENDIF

IF !Print_Choice('Ups By OrigSrc', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Ups By OrigSrc', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF

cMessage:='UpSource '+ cSrc+' '+ DTOC(dBEGDATE)+' TROUGH ' + DTOC(dENDDATE)
UPHDR(cMessage,@nPage)
DO WHILE !NEWUPS->(EOF())
  if NEWUPS->uptype=cSrc
      if NEWUPS->datein >= dbegdate
        if NEWUPS->datein <= denddate

				_PrintUprecord(cMessage,@nPage)

					//prntupline(cMess,dbegdate,dEnddate)
        endif ///begdate
      endif  //enddate
    endif /// uptype
  SKIP ALIAS NEWUPS
ENDDO
DCPRINT OFF
close all
RETURN NIL
// END INDVUPSHT


FUNCTION FRANUPSHT
LOCAL GETLIST:={} ,cProgram
LOCAL dBEGDATE,dENDDATE,cFranchise , cVehInterest , xStatus, cMessage:='' , nPage:=0
local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
LOCAL GETOPTIONS
local lMakeFile:=.F.
BOT_MAR(2)
TOP_MAR(2)
PAGE_LEN(66)
DCGETOPTIONS AUTORESIZE SAYWIDTH 0 SAYFONT '12.Courier New' getfont '12.Courier New'
cProgram:='BATUPSHT'
IF.NOT.SECURITY(@cProgram)
    DC_winALERT('SECURITY VIOLATION')
    CLOSE DATABASES
    RETURN NIL
  ENDIF
dBEGDATE:=dENDDATE:=CTOD("  /  /  ")
cVehInterest:=space(10)
cFranchise:=' '

IF !UseDb({'CLASSORT','NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
    RETURN NIL
ENDIF
XSTATUS:=.T.
@ 2,5 DCSAY 'Enter Franchise Letter' GET cFranchise ;
SAYCOLOR 3,0
@ 5,5 DCSAY 'Enter Date From' GET dBEGDATE VALID{||CHKBEGDT(dBEGDATE)}
@ 7,5 DCSAY 'Enter Date To  ' GET dENDDATE
@ 12,5 DCCHECKBOX lMakefile PROMPT 'Check Here to create Mail Merge File (upmail.dbf)'

DCREAD GUI appwindow mainwindow():drawingarea TO XSTATUS buttons 2 ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Interest Upsheets'
IF !XSTATUS
  CLOSE DATABASES
  RETURN NIL
ENDIF
if lMakeFile             //// open upmail if necessaRY
  if use_udf('upmail',.t.)
    zap
  else
    close all
  endif
endif

IF !Print_Choice('Ups By Franchise', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Ups By Franchise', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF
cMessage:='Franchise UpSheet '+ DTOC(dBEGDATE)+' TROUGH ' + DTOC(dENDDATE)
UPHDR(cMessage,@nPage)
DO WHILE !NEWUPS->(EOF())
  if NEWUPS->FRANCHISE=cFranchise
      if NEWUPS->datein >= dbegdate
        if NEWUPS->datein <= denddate
            if lMakeFile                                        /// make file if Y  upmail.dbf
             makeupmail()
            endif

				_PrintUprecord(cMessage,@nPage)

        endif ///begdate
      endif  //enddate
    endif /// franchise
  SKIP ALIAS NEWUPS
ENDDO
PRINTOFF(oPrinter)
close all
IF lMakefile
	DC_WINALERT('File UpMail.dbf has been Created.')
ENDIF

RETURN NIL


// END INDVUPSHT

FUNCTION CHKBEGDT(dBEGDATE)
SELECT NEWUPS
ORDSETFOCUS('NEWUPDT') //DBSETORDER(3)
SEEK dBEGDATE
IF !FOUND()
  DC_winALERT("NO UPS FOUND FOR BEGINING DATE ENTERED")
  RETURN .F.
ENDIF
RETURN .T.


FUNCTION BUPSHeet(cSM,dBEGDATE,dENDDATE,lMakeFile)
LOCAL GETLIST:={}, cMessage:='' , nPage:=0
BOT_MAR(2)
TOP_MAR(2)
PAGE_LEN(66)
SELECT NEWUPS
ORDSETFOCUS('NEWUPDT')
//DBSETORDER(3)
goto top
SEEK dBEGDATE
cMessage:='DAILY UPSHEET '+ DTOC(dBEGDATE)+' TROUGH ' + DTOC(dENDDATE)
UPHDR(cMessage,@nPage)
DO WHILE !NEWUPS->(EOF())
  if NEWUPS->salesman=cSM
    if NEWUPS->datein >= dbegdate
      if NEWUPS->datein <= denddate
        if lMakeFile                                        /// make file if Y  upmail.dbf
              makeupmail()
        endif

		  _PrintUprecord(cMessage,@nPage)

      endif ///begdate
    endif  //enddate
   endif  /// smi
  SKIP ALIAS NEWUPS
ENDDO
RETURN NIL

FUNCTION INBUPSHeet(cSM,dBEGDATE,dENDDATE,lMakeFile,cVehInterest)
LOCAL GETLIST:={}, cMessage:='' , nPage:=0
BOT_MAR(2)
TOP_MAR(2)
PAGE_LEN(66)
SELECT NEWUPS
ORDSETFOCUS('NEWUPDT')
//DBSETORDER(3)
goto top
SEEK dBEGDATE
cMessage:='DAILY UPSHEET for '+cSM+' '+ DTOC(dBEGDATE)+' TROUGH ' + DTOC(dENDDATE)
UPHDR(cMessage,@nPage)
DO WHILE !NEWUPS->(EOF())
  if NEWUPS->salesman=cSM
    IF NEWUPS->INTEREST=cVehInterest
    if NEWUPS->datein >= dbegdate
      if NEWUPS->datein <= denddate
        if lMakeFile                                        /// make file if Y  upmail.dbf
              makeupmail()
        endif

		  _PrintUprecord(cMessage,@nPage)


      endif ///begdate
    endif  //enddate
    ENDIF   ////cVehInterest
   endif  /// smi
  SKIP ALIAS NEWUPS
ENDDO
RETURN NIL

FUNCTION FBUPSHeet(cSM,dBEGDATE,dENDDATE,lMakeFile,cFranchise)
LOCAL GETLIST:={}  , cMessage:='' , nPage:=0
BOT_MAR(2)
TOP_MAR(2)
PAGE_LEN(66)
SELECT NEWUPS
ORDSETFOCUS('NEWUPDT')
//DBSETORDER(3)
goto top
SEEK dBEGDATE
cMessage:='DAILY UPSHEET For '+cSM+' '+ DTOC(dBEGDATE)+' TROUGH ' + DTOC(dENDDATE)
UPHDR(cMessage, @nPage)
DO WHILE !NEWUPS->(EOF())
  if NEWUPS->salesman=cSM
    IF NEWUPS->FRANCHISE=cFranchise
    if NEWUPS->datein >= dbegdate
      if NEWUPS->datein <= denddate
        if lMakeFile                                        /// make file if Y  upmail.dbf
              makeupmail()
        endif

		  _PrintUprecord(cMessage,@nPage)


      endif ///begdate
    endif  //enddate
    ENDIF   ////FRANCHISE
   endif  /// smi
  SKIP ALIAS NEWUPS
ENDDO
RETURN NIL

* -------------

function ZipAdd() // PROGRAM TO ADD AZIPCODE TO THE FILE

   LOCAL GetList[0], GetOptions
   LOCAL cZipCode, cadZone, cAreaCode, cCity
   LOCAL lStatus
   LOCAL oZip, oCOnfig

   lStatus   := .T.
   cZipcode  := Space(5)
   cCity     := Space(25)
   cAdZone   := Space(1)
   cAreaCode := Space(3)

   @   1.00,  0.00 DCSAY "ZipCode"  GET cZipcode PICTURE "XXXXX"  VALID {|| ZipSeek(cZipcode)}
   @   2.00,  0.00 DCSAY "CityState"  GET cCity
   @   3.00,  0.00 DCSAY "AdZone"  GET cAdZone
   @   4.00,  0.00 DCSAY "Area Code" GET cAreaCode

   ConfigButtons(BD_MICROSOFTBUTTONS, @oConfig, BD_SILVER)
   @   6.00,  0.00 DCPUSHBUTTONXP  CAPTION "A&dd Record" ;
                   ACCELKEY xbeK_ALT_A SIZE 12,1.3 GROUP 'BUTTONS' CONFIG oConfig TABSTOP;
                   ACTION {|| lStatus := .T., DC_ReadGuiEvent(DCGUI_EXIT_OK,GetList)} ;
                   BITMAP BITMAP_OK ;
                   TOOLTIP "Add Zip Code Record"

   @   6.00, 13.00 DCPUSHBUTTONXP CAPTION "Cancel" ;
                   ACCELKEY xbeK_ESC SIZE 12,1.3 GROUP 'BUTTONS' CONFIG oConfig TABSTOP;
                   ACTION {|| lStatus := .F., DC_ReadGuiEvent(DCGUI_EXIT_ABORT,GetList)} ;
                   BITMAP BITMAP_CANCEL1 ;
                   TOOLTIP "Abort"

   DCGETOPTIONS ;
     TITLE "Add A Zip Code" ;
     AUTORESIZE ;
     SAYWIDTH 0 ;
     SAYFONT '12.Courier New' ;
     GETFONT '12.Courier New'

   DCREAD GUI MODAL ENTEREXIT FIT OPTIONS GETOPTIONS  EVAL {|o| SetAppWindow(o)}

   IF !lStatus
     return nil
   ELSE
     oZip := ZIPPER->(DC_DbRecord():new())
     ZIPPER->(Db_Init(@oZip))

     oZip:zipcode   := cZipcode
     oZip:citystate := cCity
     oZip:adzone    := cAdZone
     oZip:area      := cAreaCode

     ZIPPER->(SaveFlds(oZip))

  ENDIF

return .T.

* -------------

function ZipSeek(cZipcode)

DbSelectArea('ZIPPER'); OrdSetFocus(1)
IF ZIPPER->(DbSeek(cZipCode))
  ALERTWIN TEXT "ZipCode Already On File" TITLE "Data Alert" TIMEOUT 5
  return .F.
ENDIF

return .T.

* -------------

*PROGRAM TO INQUIRE OR MAINTAIN A ZIPCODE

*PROGRAM TO DELETE ZIPCODE
FUNCTION DELZIP(REC)
LOCAL GETLIST:={} , getoptions, xStatus
LOCAL MCON:='Y'
DCGETOPTIONS AUTORESIZE SAYWIDTH 0
XSTATUS:=.T.
GOTO REC
@  0,  0  DCSAY "ZIPCODE"
@  0, 12  DCSAY  ZIPPER->ZIPCODE
@  1,  0  DCSAY "CITYSTATE"
@  1, 12  DCSAY  ZIPPER->CITYSTATE
@  1, 60  DCSAY "ADZONE"
@  1, 68  DCSAY  ZIPPER->ADZONE
@ 19,20 DCSAY 'IS THE ABOVE CORRECT TO DELETE Y/N' GET MCON PICTURE '@! L'
DCREAD GUI ENTEREXIT TO XSTATUS FIT OPTIONS GETOPTIONS TITLE 'DELETE A ZIP CODE'
IF !XSTATUS
  RETURN NIL
ENDIF

IF MCON = "N"
  RETURN NIL
ENDIF
IF MCON = "Y"
  IF ZIPPER->(DC_RECLOCK())
    DBDELETE(ZIPPER->(RECNO()))
  ENDIF
  ZIPPER->(DBCOMMIT())
  ZIPPER->(DBRUNLOCK())
	PACK
	ZIPPER->(DBCOMMIT())
ENDIF
RETURN NIL

// END DELZIP

*PROGRAM TO ADD AN ADSOURCE
FUNCTION ADADD()
LOCAL GETLIST:={} , xstatus , getoptions  , cAds ,cDes
LOCAL clMCON:='Y'
DCGETOPTIONS AUTORESIZE SAYWIDTH 0  sayfont '10.Courier New' getfont '10.Courier New'

IF !UseDb({'ADSOURCE'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF

cADS:=space(4)
cDes:= SPACE(10)
@  10,  0  DCSAY "Enter Advertising Source"  Get  cAds Picture "XXXX"
@  11,  0  DCSAY "Enter Description"  GET  cDes PICTURE "XXXXXXXXXX"
@ 12,0 DCSAY 'OK to Update Esc=Exit No Update' get clMcon picture '@! L'
DCREAD GUI MODAL ENTEREXIT TO XSTATUS FIT ADDBUTTONS OPTIONS GETOPTIONS TITLE 'ADD ADVERTISING SOURCE' EVAL{|o| setappwindow(o)}
IF !XSTATUS
  RETURN NIL
ENDIF

SEEK cADS
IF FOUND()
  DC_winALERT( 'AD SOURCE ALREADY EXISTS')
  RETURN NIL
ENDIF

IF !XSTATUS
  RETURN NIL
ENDIF
IF clMcon='N'
	return nil
ENDIF
IF clMcon='Y'
 IF ADSOURCE->(DC_ADDREC(3,.T.,2))
	REPLACE ADSOURCE->ADSRCE WITH cADS
	REPLACE ADSOURCE->DESC WITH cDES
 ENDIF
endif
ADSOURCE->(DBCOMMIT())
CLOSE ALL
RETURN .T.

// END ADADD


*PROGRAM TO ADD A SALESPERSON
FUNCTION SMADD()
LOCAL GETLIST:={} , getoptions , xStatus
LOCAL cSM:= "  " , cName , cNum
LOCAL MCON:='Y'
DCGETOPTIONS AUTORESIZE SAYWIDTH 0
XSTATUS:=.T.
cNaME := SPACE(25)
cSM := SPACE(2)
cNUM := SPACE(3)
/// put up dialogue

@ 1,0  DCSAY "Enter Salesperson Initials    "   GET  cSM PICTURE "@!" VALID{|o| _isSMexist(cSm)}
@ 2,0  DCSAY "Enter Salesperson Name        "  GET  cNaME
@ 3,0  DCSAY 'Enter Salesperson Employee Id '  GET cNUM VALID{||CHKSMNUM(cNUM)}
@ 4,50 DCSAY "IS THIS CORRECT Y/N"  GET MCON PICTURE "@! L"
DCREAD GUI MODAL ENTEREXIT TO XSTATUS addbuttons FIT OPTIONS GETOPTIONS TITLE 'ADD A SALESPERSON' EVAL{|o| Setappwindow(o)}
IF !XSTATUS
  RETURN NIL
ENDIF

IF MCON = "Y"
	IF sm->(dc_addrec(3,.T.,2))
		REPLACE SM->SALESMAN WITH cSM
		REPLACE SM->NAME WITH cNaME
		REPLACE SM->ID WITH cNUM
		REPLACE SM->ACTIVE with 'Y'
		UNLOCK
	ENDIF
ENDIF
RETURN nil

// END SMADD.PRG

FUNCTION ADsourceadd()
LOCAL GETLIST:={} , getoptions , xStatus
LOCAL cSource:=space(4) , cName:=space(10)
LOCAL MCON:='Y'
DCGETOPTIONS AUTORESIZE SAYWIDTH 0
XSTATUS:=.T.
/// put up dialogue

@ 1,0  DCSAY "Enter Advertising Source Code    "   GET  cSource PICTURE "@!" VALID{|o| _isSrcexist(cSource)}
@ 2,0  DCSAY "Enter Source Description         "  GET  cNaME
@ 4,50 DCSAY "IS THIS CORRECT Y/N"  GET MCON PICTURE "@! L"
DCREAD GUI MODAL ENTEREXIT TO XSTATUS addbuttons FIT OPTIONS GETOPTIONS TITLE 'ADD A SALESPERSON' EVAL{|o| Setappwindow(o)}
IF !XSTATUS
  RETURN NIL
ENDIF

IF MCON = "Y"
	IF ADSOURCE->(dc_addrec(3,.T.,2))
		REPLACE ADSOURCE->ADSRCE WITH cSource
		REPLACE ADSOURCE->DESC WITH cName
		UNLOCK
	ENDIF
ENDIF
RETURN nil

// END adsourceADD.PRG



static function _isSmExist(cSM)
  SELECT SM
  SEEK cSM
  IF FOUND()
   DC_winALERT('SALESPERSON ALREADY EXISTS')
   RETURN FALSE
  ENDIF
RETURN TRUE

static function _isSrcexist(cSource)
  SELECT ADSOURCE
  SEEK cSource
  IF FOUND()
   DC_winALERT('Source Already Exists')
   RETURN FALSE
  ENDIF
  RETURN TRUE



FUNCTION CHKSMNUM(cNUM)       //// MAKE SURE FIELD IS POPULATED
  IF EMPTY(cNUM)
    RETURN .F.
  ENDIF
  RETURN .T.

*PROGRAM TO CHANGE A SALESPERSON
FUNCTION SMSCREEN()
local getlist:={}
local getoptions  , xStatus ,  clActive
LOCAL cSM:=SPACE(2)
local clok:='Y', cName,cNum
DCGEToptions saywidth 0 sayfont '12.Courier New' Getfont '12.Courier New'   autoresize

 IF !UseDb({'SM'})
 	 DC_WINALERT('Cannot Open Files     .')
    	 RETURN NIL
 ENDIF


@ 10,10 DCSAY 'ENTER INITIALS TO CHANGE'  GET cSM PICTURE "!!"
dcREAD gui modal enterexit to xstatus fit options getoptions title 'Sales Person Maintain' eval{|o| setappwindow(o)}

SELECT SM
SEEK cSM
IF .NOT. FOUND()
  dc_winalert( 'SALESPERSON DOES NOT EXIST')
	RETURN .T.
ENDIF
cName :=SM-> NAME
cSM :=SM-> SALESMAN
cNUM:=SM-> ID
clActive:=SM->active
@  5,0  DCSAY "ENTER SALESPERSON' INITIALS"
@  5,41  DCGET  cSM
@  7,0  DCSAY "ENTER SALESPERSON'S NAME"
@  7,41  DCGET  cNaME PICTURE "XXXXXXXXXXXXXXXXXXXXXXXXXX"
@ 12,0 DCSAY 'ENTER SALESPERSON ID'
@ 12,41 DCGET cNUM PICTURE "XXX"
@ 13,0 DCSAY 'Active Y/N' get clActive picture '@! L'
@ 15,0 DCSAY 'IS THE ABOVE OK Y/N'
@ 15,30 DCGET clOK PICTURE '@! L'
dcREAD gui enterexit fit options getoptions title "Sales Person Screen"

IF clOK ='Y'
  IF sm->(dc_reclock())
		REPLACE sm->SALESMAN WITH cSM
		REPLACE sm->NAME WITH cNaME
		REPLACE sm->ID WITH cNUM
		replace sm->active with clActive
		UNLOCK
	ENDIF
ENDIF &&OK=Y
RETURN nil

// END SMSCREEN

*PROGRAM TO DELETE A SALESPERSON
FUNCTION SMDELETE(nRec)
local getlist:={} , xstatus
local getoptions
LOCAL cSM,cNaME,mcon
DCGEToptions saywidth 0 sayfont '12.Courier New' Getfont '12.Courier New'   autoresize
SELECT SM
SM->(DBGOTO(nRec))
mcon:='Y'
cSM:=SM->SALESMAN
cNaME:=SM->NAME
@  5,0  DCSAY "Salesperson' Initials"
@  5,41  DCSAY  cSM
@  7,0  DCSAY "Salesperson'S Name   "
@  7,41  DCSAY  cName PICTURE "XXXXXXXXXXXXXXXXXXXXXXXXXX"
@ 16,50 DCSAY 'Is This The Record To Delete Y/N'  GET MCON PICTURE "@! L"
dcREAD gui modal enterexit fit options getoptions title "Sales Person Delete Screen" ;
eval{|o| setappwindow(o)}

IF MCON = "Y"
	SELECT SM
	bv_blank(.f.,.t.)
ENDIF
RETURN .T.

// END SMDELETE
FUNCTION ADSOURCEDELETE(nRec)
local getlist:={} , xstatus
local getoptions
LOCAL cAdSource,cNaME,mcon
DCGEToptions saywidth 0 sayfont '12.Courier New' Getfont '12.Courier New'   autoresize
SELECT ADSOURCE
ADSOURCE->(DBGOTO(nRec))
mcon:='Y'
caDsOURCE:=ADSOURCE->ADSRCE
cNaME:=ADSOURCE->DESC
@  5,0  DCSAY "Advertising Source"
@  5,41  DCSAY  cAdsource
@  7,0  DCSAY "Description   "
@  7,41  DCSAY  cName PICTURE "XXXXXXXXXXXXXXXXXXXXXXXXXX"
@ 16,50 DCSAY 'Is This The Record To Delete Y/N'  GET MCON PICTURE "@! L"
dcREAD gui modal enterexit fit options getoptions title "Advertising Source Delete Screen" ;
eval{|o| setappwindow(o)}

IF MCON = "Y"
	SELECT ADSOURCE
	bv_blank(.f.,.t.)
ENDIF
RETURN .T.

// END ADDSRCDELETE



FUNCTION LISTZIP()
local getlist:={} , xStatus
local nCopies:=1,nOrient:=1,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'

BOT_MAR(2)
TOP_MAR(2)
PAGE_LEN(66)

IF !Print_Choice('Local Zipcode List', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Local Zipcode List', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF

select zipper
goto top
@ dc_printerrow()+1,0 DCPRINT SAY 'Up System ZipCode Printout as of      ' + dtoc(date())
@ dc_printerrow()+1,0 DCPRINT SAY 'ZipCode  Citystate       Adzone'

do while !zipper->(eof())
  @ dc_printerrow()+1,0 DCPRINT SAY zipper->zipcode
  @ dc_printerrow(),10  DCPRINT SAY zipper->citystate
  @ dc_printerrow(),35  DCPRINT SAY zipper->adzone
  if dcpageeject()
    @ dc_printerrow()+1,0 DCPRINT SAY 'ZipCode  Citystate             Adzone'
  endif
  skip alias zipper
enddo
printoff(oPrinter)
zipper->(dbgotop())
RETURN .T.

// END LISTZIP




FUNCTION ZIPCOUNT()
LOCAL GETLIST:={}  , nAlltotal , nZonetotal, cZone , cZipcode , nZiptotal
LOCAL GETOPTIONS, xStatus , clSkipads
local dbegdate,denddate
local nCopies:=1,nOrient:=1,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New',nPage
DCGETOPTIONS SAYWIDTH 0 AUTORESIZE SAYFONT '12.Courier New' getfont '12.Courier New'

IF !UseDb({'ZIPPER','NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF

dbegdate:=denddate:=ctod('  /  /  ')


TOP_MAR(3)
BOT_MAR(2)
PAGE_LEN(62)
nPAGE:=0

clskipads:='Y'
@ 5,0 DCSAY 'Print Ups Report by Zip code'
@ 10,0 DCSAY 'Enter Beginning date to count ups for' get dbegdate
@ 12,0 DCSAY 'Enter Ending date to count ups for   ' get denddate
@ 13,0 DCSAY 'Do you wish to skip Blank or Zero AdZones' get clskipads
DCREAD GUI appwindow mainwindow():drawingarea TO XSTATUS buttons 2 ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'ZipCode Count'

if !xstatus
  return nil
endif
IF !Print_Choice('Zipcode Ups', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn(' Zipcode Ups', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF

@ dc_printerrow(),1 DCPRINT SAY DATE()
@ dc_printerrow(),dc_printerCOL()+4 DCPRINT SAY TIME()
@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY 'Zipcount'
@ dc_printerrow()+1,0 DCPRINT SAY dBegdate
@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY 'thru'
@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY dEnddate

select zipper
goto top
if clskipads='Y'
  do while empty(zipper->adzone)
    skip alias zipper
  enddo
endif


nAlltotal:=0
do while !zipper->(eof())
  select newups
  seek dbegdate
  nZonetotal:=0
  cZone:=zipper->adzone
   do while zipper->adzone=czone
       czipcode:=zipper->zipcode
       nziptotal:=0
       select newups
       goto top
       seek dbegdate
       do while !NEWUPS->(eof())
          if NEWUPS->datein >=dbegdate
            if NEWUPS->datein <=denddate
             if NEWUPS->zipcode=czipcode
              nZiptotal:=nZiptotal+1
              nzonetotal:=nzonetotal+1
              nalltotal:=nalltotal+1
             endif
            endif
          endif
          skip alias newups
       enddo
       if nziptotal > 0             /// only print if > 0
        select zipper
        @ dc_printerrow()+1,0 DCPRINT SAY zipper->zipcode
        @ dc_printerrow(),20 DCPRINT SAY zipper->citystate
        @ dc_printerrow(),60 DCPRINT SAY 'Upcount ' + str(nziptotal)
        if dcpageeject()
          @ dc_printerrow()+1,0 DCPRINT SAY 'Zipcode Report Continued'
        endif
       endif
   skip alias zipper
  enddo
  @ dc_printerrow()+2,0 DCPRINT SAY 'Zone '+czone                font '10.Courier New Bold'
  @ dc_printerrow(),20 DCPRINT SAY 'Zone Total '+str(nzonetotal)   font '10.Courier New Bold'
  @ dc_printerrow()+1,20 DCPRINT SAY ' '
  if dcpageeject()
       @ dc_printerrow()+1,0 DCPRINT SAY 'Zipcode Report Continued'
  endif
enddo
@ dc_printerrow(),20 DCPRINT SAY 'Report Total '+str(nalltotal)
if dcpageeject()
      @ dc_printerrow()+1,0 DCPRINT SAY 'Zipcode Report Continued'
endif
printoff(oPrinter)
close all
return nil

FUNCTION INTCOUNT()       /// count ups by interest
LOCAL GETLIST:={}
LOCAL GETOPTIONS ,  nAlltotal, nClasstotal, cClass,  xStatus
local dbegdate,denddate,nPage
local nCopies:=1,nOrient:=1,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
DCGETOPTIONS SAYWIDTH 0 AUTORESIZE SAYFONT '12.Courier New' getfont '12.Courier New'

IF !UseDb({'CLASSORT','NEWUPS',''})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF

NEWUPS->(ORDSETFOCUS('NEWUPDT'))

TOP_MAR(3)
BOT_MAR(2)
PAGE_LEN(62)
nPAGE:=0

dbegdate:=denddate:=ctod('  /  /  ')

@ 5,0 DCSAY 'Print ups Report by Interest'
@ 10,0 DCSAY 'Enter Beginning date to count ups for' get dbegdate
@ 12,0 DCSAY 'Enter Ending date to count ups for   ' get denddate
dcread gui appwindow mainwindow():drawingarea enterexit options getoptions to xstatus title 'Automan Ups System Interest Count' fit
if !xstatus
  close all
  return nil
endif

IF !Print_Choice('Interest Ups', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn(' Interest Ups', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF
@ dc_printerrow(),1 DCPRINT SAY DATE()
@ dc_printerrow(),dc_printerCOL()+4 DCPRINT SAY TIME()
@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY 'Intcount'
@ dc_printerrow()+1,0 DCPRINT SAY dBegdate
@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY 'thru'
@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY dEnddate

select classort
goto top
nalltotal:=0
do while !classort->(eof())
  cClass:=classort->classtype
  select newups
  seek dbegdate
  nClasstotal:=0
  do while !NEWUPS->(eof())
        if NEWUPS->datein >=dbegdate
          if NEWUPS->datein <=denddate
            if NEWUPS->interest=cClass
             nclasstotal:=nclasstotal+1
             nalltotal:=nalltotal+1
            endif
          endif
        endif
        skip alias newups
   enddo
   if nclasstotal > 0             /// only print if > 0
       select classort
       @ dc_printerrow()+1,0 DCPRINT SAY classort->classtype
       @ dc_printerrow(),60 DCPRINT SAY 'Upcount ' + str(nclasstotal)
       if dcpageeject()
        @ dc_printerrow()+1,0 DCPRINT SAY 'Interest Report Continued'
       endif
    endif
   skip alias classort
enddo
@ dc_printerrow()+1,20 DCPRINT SAY 'Report Total '+str(nalltotal)
if dcpageeject()
      @ dc_printerrow()+1,0 DCPRINT SAY 'Interest Report Continued'
endif
printoff(oPrinter)
close all
return nil

FUNCTION AdSourceCount()       /// count ups by adsource
LOCAL GETLIST:={}
LOCAL GETOPTIONS , nAlltotal, nAdsTotal, cAdsrce,  xStatus  , nSold  ,nClasstotal
local dBegdate,dEnddate
local nCopies:=1,nOrient:=1,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New' ,nPage
DCGETOPTIONS SAYWIDTH 0 AUTORESIZE SAYFONT '12.Courier New' getfont '12.Courier New'

IF !UseDb({'ADSOURCE','NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF

NEWUPS->(ORDSETFOCUS('NEWUPDT'))

TOP_MAR(2)
BOT_MAR(2)
PAGE_LEN(62)
nPAGE:=0

dBegdate:=dEnddate:=ctod('  /  /  ')

@ 5,0 DCSAY 'Print ups Report by Advertising  Source'
@ 10,0 DCSAY 'Enter Beginning date to count ups for' get dBegdate
@ 12,0 DCSAY 'Enter Ending date to count ups for   ' get dEnddate
dcread gui appwindow mainwindow():drawingarea enterexit options getoptions to xstatus title 'Automan Ups System Interest Count' fit
if !xstatus
  close all
  return nil
endif

IF !Print_Choice('AdSource Ups', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn(' AdSource Ups', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF

@ dc_printerrow(),1 DCPRINT SAY DATE()
@ dc_printerrow(),dc_printerCOL()+4 DCPRINT SAY TIME()
@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY 'AdvSourceCount'
@ dc_printerrow()+1,0 DCPRINT SAY dBegdate
@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY 'thru'
@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY dEnddate

select ADSOURCE
goto top
nalltotal:=0
do while !ADSOURCE->(eof())
  cAdsrce:=ADSOURCE->ADSRCE
  select newups
  seek dbegdate
  nClasstotal:=0
  nSold:=0
  do while !NEWUPS->(eof())
        if NEWUPS->datein >=dbegdate
          if NEWUPS->datein <=denddate
            if NEWUPS->ADSOURCE=cAdsrce
             nclasstotal:=nclasstotal+1
             nalltotal:=nalltotal+1
				 IF NEWUPS->UPSTATUS='S'
					nSold++
				 ENDIF
            endif
          endif
        endif
        skip alias newups
   enddo
   if nclasstotal > 0             /// only print if > 0
       select Adsource
       @ dc_printerrow()+1,0 DCPRINT SAY Adsource->desc
       @ dc_printerrow(),60 DCPRINT SAY 'Upcount ' + alltrim(str(nclasstotal))
		 @ dc_printerrow(),80 DCPRINT SAY 'Sold '+ alltrim(str(nSold))
		 @ dc_printerrow(),92 DCPRINT SAY nSold/nClasstotal picture '99.99'
       if dcpageeject()
        @ dc_printerrow()+1,0 DCPRINT SAY 'Adsource Report Continued'
       endif
    endif
   skip alias ADSOURCE
enddo
@ dc_printerrow()+1,20 DCPRINT SAY 'Report Total '+str(nalltotal)
if dcpageeject()
      @ dc_printerrow()+1,0 DCPRINT SAY 'Adsource Report Continued'
endif
printoff(oPrinter)
close all
return nil





FUNCTION ZIPcodeCOUNT()
LOCAL GETLIST:={} , oDialog
LOCAL GETOPTIONS , xstatus, cZipcode , cZipdesc, nAlltotal, nZiptotal,nPage
local nCopies:=1,nOrient:=1,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
local dbegdate,denddate
DCGETOPTIONS SAYWIDTH 0 AUTORESIZE SAYFONT '12.Courier New' getfont '12.Courier New'

dbegdate:=denddate:=ctod('  /  /  ')

 IF !UseDb({'NEWUPS','ZIPPER','USAZIP'})
 	 DC_WINALERT('Cannot Open Files     .')
    	 RETURN NIL
 ENDIF

 NEWUPS->(ORDSETFOCUS('NEWUPZIP'))


TOP_MAR(2)
BOT_MAR(2)
PAGE_LEN(62)
nPAGE:=0

@ 5,0 DCSAY 'Print ups Report by Zip code'
@ 10,0 DCSAY 'Enter Beginning date to count ups for' get dbegdate
@ 12,0 DCSAY 'Enter Ending date to count ups for   ' get denddate
dcread gui appwindow mainwindow():drawingarea enterexit options getoptions to xstatus title 'Automan Ups System Zip Count' fit
if !xstatus
  close all
  return nil
endif


IF !Print_Choice('ZipCode Ups', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn(' ZipCode Ups', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF


@ dc_printerrow(),1 DCPRINT SAY DATE()
@ dc_printerrow(),dc_printerCOL()+4 DCPRINT SAY TIME()
@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY 'ZipCodeCount'
@ dc_printerrow()+1,0 DCPRINT SAY dbegdate
@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY 'thru'
@ dc_printerrow(),dc_printerCOL()+8 DCPRINT SAY denddate
nalltotal:=0
select newups
do while !NEWUPS->(eof())
        czipcode:=NEWUPS->zipcode
        nziptotal:=0
        czipdesc:=' '
        do while NEWUPS->zipcode=czipcode
          if NEWUPS->datein >=dbegdate
            if NEWUPS->datein <=denddate
              nZiptotal:=nziptotal+1
              nalltotal:=nalltotal+1
             endif
          endif
          skip alias newups
        enddo
        select zipper
        seek czipcode
        if found()
          czipdesc:=zipper->citystate
        else
          addtozip(czipcode,@czipdesc)
        endif

        if nziptotal >0
          @ dc_printerrow()+1,0 DCPRINT SAY czipcode
          @ dc_printerrow(),12 DCPRINT SAY czipdesc
          @ dc_printerrow(),60 DCPRINT SAY 'Upcount ' + str(nziptotal)
          if dcpageeject()
            @ dc_printerrow()+1,0 DCPRINT SAY 'Zipcode Report Continued'
          endif
        endif
  enddo
  @ dc_printerrow()+2,20 DCPRINT SAY 'Report Total '+str(nalltotal)
  if dcpageeject()
       @ dc_printerrow()+1,0 DCPRINT SAY 'Zipcode Report Continued'
  endif

printoff(oPrinter)
close all
return nil

static function addtozip(cZipcode,cZipdesc)
local getlist:={}
local claddtoz:='Y' , getoptions , xStatus
DCGEToptions saywidth 0 autoresize sayfont '12.Courier New' getfont '12.Courier New'
select usazip
seek cZipcode
if found()
  cZipdesc:=usazip->city+','+usazip->state
  @ 5,0 DCSAY usazip->city+' '+usazip->state
  @ 10,0 DCSAY 'Add This Zipcode to the Master File of ZipCodes YOU WANT TO TRACK Y/N' get clAddtoz
  dcread gui to xstatus enterexit fit options getoptions title 'Add Zipcode to master'
  if clAddtoz='Y'
    select zipper
    if zipper->(dc_addrec())
      replace zipper->zipcode with usazip->zip
      replace zipper->citystate with usazip->city+','+usazip->state
      replace zipper->area with usazip->areacode
      ZIPPER->(dbcommit())
    endif
  endif
endif
return nil


FUNCTION LISTADS
local getlist:={} , lPrevue
local nCopies:=1,nOrient:=1,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
BOT_MAR(2)
TOP_MAR(2)
PAGE_LEN(66)
IF !Print_Choice('Advertising Source List', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Ad Source List', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF

select adsource
goto top
@ dc_printerrow()+1,0 DCPRINT SAY 'Up System Ad Source Printout as of      ' + dtoc(date())
@ dc_printerrow()+1,0 DCPRINT SAY 'Ad Source  Description'

do while !adsource->(eof())
  @ dc_printerrow()+1,0 DCPRINT SAY adsource->adsrce
  @ dc_printerrow(),10  DCPRINT SAY adsource->desc
  if dcpageeject()
    @ dc_printerrow()+1,0 DCPRINT SAY 'Ad Source Description'
  endif

  skip alias adsource
enddo
printoff(oPrinter)
ADSOURCE->(DBGOTOP())
RETURN .T.


// END LISTADS

FUNCTION SMLIST()
local getlist:={}
local nCopies:=1,nOrient:=1,nDefmode:=4, oPrinter,cOutfile:='', cFont:='10.Courier New'
BOT_MAR(2)
TOP_MAR(2)
PAGE_LEN(66)
IF !Print_Choice('Sales Person List', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Sales Person List', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF
select sm
goto top
@ dc_printerrow()+1,0 DCPRINT SAY 'Up System Sales Persons      ' + dtoc(date())
@ dc_printerrow()+1,0 DCPRINT SAY 'SalesInit     ID      Name'

do while !sm->(eof())
  @ dc_printerrow()+1,0 DCPRINT SAY sm->salesman
  @ dc_printerrow(),10  DCPRINT SAY sm->id
  @ dc_printerrow(),20  DCPRINT SAY sm->name
  @ dc_printerrow(),40  DCPRINT SAY sm->ACTIVE
  if dcpageeject()
    @ dc_printerrow()+1,0 DCPRINT SAY 'SalesInit'
	 @ dc_printerrow(),10 DCPRINT SAY  'Emp ID# '
	 @ dc_printerrow(),20 DCPRINT SAY  'Name  '
	 @ dc_printerrow(),40 DCPRINT SAY ' Active'
  endif

  skip alias sm
enddo
printoff(oPrinter)
SM->(DBGOTOP())
RETURN .T.

// END SMLIST

*COPYUPS.PRG PROGRAM TO COPY NEW UPS TO HISTROY FILE
FUNCTION COPYUPS()
LOCAL GETLIST:={} , cProgram , getoptions, xstatus, oDialog
LOCAL dDATEBACK:=CTOD("  /  /  ") ,nTotrec , clVercon ,oRecord
cProgram:='COPYUPS'
IF.NOT.SECURITY(@cProgram)
    DC_WINALERT('SECURITY VIOLATION')
  ENDIF

DCGETOPTIONS AUTORESIZE SAYWIDTH 0
XSTATUS:=.T.
cProgram:="COPYUPS"
IF .NOT.SECURITY(@cProgram)
  DC_WINALERT('SECURITY VIOLATION')
  RETURN NIL
ENDIF



IF !UseDb({'NEWUPS','HNEWUPS'},1,.F.)
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF

nTOTREC:=NEWUPS->(LASTREC())

@ 2,0 DCSAY 'ENTER DATE TO COPY BACK FROM '  GET dDATEBACK
DCREAD GUI ENTEREXIT TO XSTATUS FIT OPTIONS GETOPTIONS TITLE 'COPY UPS TO HISTORY'
IF !XSTATUS
  RETURN NIL
ENDIF
clVERCON:='Y'
@ 3,0 DCSAY 'YOU ARE ABOUT TO COPY  ALL UP RECORDS THAT HAVE A DATE IN BEFORE'
@ 3,70 DCSAY dDATEBACK
@ 4,0 DCSAY 'IS THIS WHAT YOU WANT Y /N OR ESC=EXIT'  GET clVERCON PICTURE '@! L'
DCREAD GUI ENTEREXIT TO XSTATUS FIT OPTIONS GETOPTIONS TITLE 'ADD A SALESPERSON'
IF !XSTATUS
  RETURN NIL
ENDIF

IF clVERCON='N'
  RETURN NIL
ENDIF

SELECT NEWUPS
oDIALOG:=DC_WAITON("NOW COPYING UP RECORDS")

DO WHILE !NEWUPS->(EOF())

  IF NEWUPS->DATEIN < dDateback
	IF !NEWUPS->(DELETED())
		oRecord:=NEWUPS->(DC_DBRECORD():NEW())
		NEWUPS->(DC_DBSCATTER(oRecord))
		HNEWUPS->(DC_DBGATHER(oRecord,.t.))
		NEWUPS->(DBDELETE())
	ENDIF
  ENDIF
  SELECT NEWUPS
  NEWUPS->(DBSKIP())
ENDDO

//CLOSE DATABASES
SELECT NEWUPS
CLOSE INDEXES
NEWUPS->(DBPACK())
NEWUPS->(DBCOMMIT())
DC_IMPL(oDIALOG)
CLOSE DATABASES

IF !UseDb({'NEWUPS'},1,.F.,.T.)
	 DC_WINALERT('Cannot Open Up System Files     .')
    RETURN NIL
ENDIF

DC_WINALERT('UP COPY NOW COMPLETE- PRESS ANY KEY')

CLOSE ALL

RETURN NIL
// END PURGUPS

*UPSTAT
FUNCTION CLOSEREPT()
LOCAL GETLIST:={} , getoptions, cFranchise ,xStatus , dBegdate, dEnddate
local nU1,nU2,nU3,nU4,nU5,nU6,nS1,nS2,nS3,nS4,nS5,nS6,nCP1,nCP2,nCP3,nCP4,nCP5,nCP6,nNEW,nNEWS,nUSED,nUSEDS,nDDN,nDDU,nDemo,nTOmgr,;
	nPupsIn,nIupsIn,aSoldarray:={} , aUpsInArray:={} ,lAddedAsSold:=.f.

IF !UseDb({'NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF



DCGETOPTIONS SAYWIDTH 0 AUTORESIZE sayfont '12.Courier New' getfont '12.Courier New'
cFranchise:=SPACE(2)
nU1:=nU2:=nU3:=nU4:=nU5:=nU6:=nS1:=nS2:=nS3:=nS4:=nS5:=nS6:=nCP1:=nCP2:=nCP3:=nCP4:=nCP5:=nCP6:=nNEW;
	:=nNEWS:=nUSED:=nUSEDS:=nDDN:=nDDU:=nDemo:=nTOmgr:=nPupsIn:=nIupsIn:=0
dBEGDATE:=dENDDATE:=CTOD("  /  /  ")
@ 10,30 DCSAY 'Enter Beginning Date'  GET dBEGDATE valid{|| chkBEGDT(dBEGDATE)} SAYFONT '12.COURIER' GETFONT '12.COURIER'
@ 11,30 DCSAY 'Enter Ending Date   ' GET dENDDATE SAYFONT '12.COURIER'  GETFONT '12.COURIER'
@ 12,30 DCSAY 'Enter Franchise Or Leave Blank For All' GET cFranchise
DCREAD GUI modal fit ENTEREXIT BUTTONS 2 TO XSTATUS options getoptions  TITLE 'UPCOUNT Report';
eval{|o| setappwindow(o)}
IF !XSTATUS
  RETURN NIL
ENDIF
if !empty(cFranchise)
  set filter to NEWUPS->franchise=cFranchise
endif
SELECT newups
GOTO TOP
ORDSETFOCUS('NEWUPDT')
seek dbegdate
DO WHILE !NEWUPS->(EOF())
  IF NEWUPS->DATEIN=NEWUPS->BEBACK3 /// BEBACK3 IS REALLY THE ORIGINAL DATE IN ONLY COUNT AS A 1OR23OR4 IF THESE DATES ARE THE SAME
   IF NEWUPS->DATEIN >=dBEGDATE
     IF NEWUPS->DATEIN <=dENDDATE
		IF NEWUPS->DATEIN <> NEWUPS->BEBACK1       /// DO NOT COUNT AN UP IF THE DATEIN AND BEBACK DATES ARE THE SAME
			IF NEWUPS->DATEIN <> NEWUPS->BEBACK2    /// THEY WILL BE COUNTED AS BEBACKS
        		IF NEWUPS->UPTYPE = 1
        		        nU1:=nU1+1
						  AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 1'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

		  					 IF NEWUPS->UPSTATUS='S'
								IF EMPTY(NEWUPS->BEBACK1)
									IF EMPTY(NEWUPS->BEBACK2)
		  						    nS1:=nS1+1
								    AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 1'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
		  					      ENDIF
							   ENDIF
							 ENDIF
        		ENDIF
        		IF NEWUPS->UPTYPE = 2
        		        nU2:=nU2+1
						  AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 2'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

		  					 IF NEWUPS->UPSTATUS='S'
								IF EMPTY(NEWUPS->BEBACK1)
									IF EMPTY(NEWUPS->BEBACK2)
		  						     nS2:=nS2+1
								     AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 2'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
									ENDIF
								ENDIF
		  					 ENDIF

        		ENDIF
        		IF NEWUPS->UPTYPE = 3
        		        nU3:=nU3+1
						  AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

		  					 IF NEWUPS->UPSTATUS='S'
								IF EMPTY(NEWUPS->BEBACK1)
									IF EMPTY(NEWUPS->BEBACK2)
		  						     nS3:=nS3+1
								     AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
									ENDIF
								ENDIF
		  					 ENDIF

        		ENDIF
        		IF NEWUPS->UPTYPE = 4
        		        nU4:=nU4+1
						  AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

		  					 IF NEWUPS->UPSTATUS='S'
								IF EMPTY(NEWUPS->BEBACK1)
									IF EMPTY(NEWUPS->BEBACK2)
		  						     nS4:=nS4+1
								     AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
		  					      ENDIF
								ENDIF
							 ENDIF
        		ENDIF
			ENDIF //// <> BEBACK1
		ENDIF /// <> BEBACK2






		IF NEWUPS->ORIGSOURCE = 'I'
                nU5:=nU5+1
					 IF NEWUPS->UPSTATUS='S'
						nS5:=nS5+1
					 ENDIF
					 IF NEWUPS->UPTYPE < 5
						nIupsIn++
					 ENDIF

      ENDIF
		IF NEWUPS->ORIGSOURCE = 'P'
                nU6:=nU6+1
					 IF NEWUPS->UPSTATUS='S'
						nS6:=nS6+1
					 ENDIF
					 IF NEWUPS->UPTYPE < 5
						nPupsIn++
					 ENDIF
      ENDIF

      IF NEWUPS->NEWUSED='N'
          nNEW:=nNEW+1
          IF NEWUPS->upstatus='S'
            nNEWS:=nNEWS+1
          ENDIF
          IF NEWUPS->DOWNDEAL > 0
             nDDN:=nDDN+1
          ENDIF
      ENDIF
      IF NEWUPS->NEWUSED='U'
          nUSED:=nUSED+1
          IF NEWUPS->upstatus='S'
            nUSEDS:=nUSEDS+1
          ENDIF
          IF NEWUPS->DOWNDEAL > 0
             nDDU:=nDDU+1
          ENDIF
      ENDIF
		if NEWUPS->demo='Y'
					nDemo++
		endif
		if !empty(NEWUPS->tomgr)
					nTOmgr++
		endif
	  ENDIF /// DATEIN
   ENDIF   /// DATEIN
  ENDIF /// BEBACK3 WHICH IS REALLY THE ORIGDATEIN

  /// IF ORIGINAL DATE IN IS IN RANGE BUT NOT EQUAL TO LAST DATE IN
  IF NEWUPS->DATEIN > NEWUPS->BEBACK3 /// BEBACK3 IS REALLY THE ORIGINAL DATE IN ONLY COUNT AS A 1OR23OR4 IF THESE DATES ARE THE SAME
	   IF NEWUPS->BEBACK3 >=dBEGDATE
         IF NEWUPS->BEBACK3 <=dENDDATE
	        IF NEWUPS->BEBACK3 <> NEWUPS->BEBACK1       /// DO NOT COUNT AN UP IF THE DATEIN AND BEBACK DATES ARE THE SAME
			     IF NEWUPS->BEBACK3 <> NEWUPS->BEBACK2
					  IF NEWUPS->UPTYPE = 1
        		        nU1:=nU1+1
						  AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 1'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

		  					 IF NEWUPS->UPSTATUS='S'
								IF EMPTY(NEWUPS->BEBACK1)
									IF EMPTY(NEWUPS->BEBACK2)
		  						       nS1:=nS1+1
								       AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 1'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
							      ENDIF
								ENDIF
							 ENDIF
         	   	ENDIF
        		      IF NEWUPS->UPTYPE = 2
        		        nU2:=nU2+1
						  AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 2'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

		  					 IF NEWUPS->UPSTATUS='S'
								IF EMPTY(NEWUPS->BEBACK1)
									IF EMPTY(NEWUPS->BEBACK2)
		  						       nS2:=nS2+1
								       AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 2'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
							      ENDIF
								ENDIF
							 ENDIF
        		      ENDIF
        		      IF NEWUPS->UPTYPE = 3
        		        nU3:=nU3+1
						  AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

		  					 IF NEWUPS->UPSTATUS='S'
								IF EMPTY(NEWUPS->BEBACK1)
									IF EMPTY(NEWUPS->BEBACK2)
		  						     nS3:=nS3+1
								     AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
									ENDIF
								ENDIF
		  					 ENDIF

        		      ENDIF
        		      IF NEWUPS->UPTYPE = 4
        		        nU4:=nU4+1
						  AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

		  					 IF NEWUPS->UPSTATUS='S'
								IF EMPTY(NEWUPS->BEBACK1)
									IF EMPTY(NEWUPS->BEBACK2)
		  						     nS4:=nS4+1
								     AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
		  					      ENDIF
								ENDIF
							 ENDIF
        		      ENDIF
					 ENDIF /// NOT A BEBACK1
				 ENDIF /// NOT A BEBACK2
			 ENDIF  // < ENDDATE
	  	ENDIF  /// > BEGDATE
  ENDIF  /// ORIG DATE IN NOT EQUAL TO LAST DATE IN





  SKIP ALIAS NEWUPS
ENDDO
SELECT NEWUPS
GOTO BOTTOM
SKIP -1000
DO WHILE !NEWUPS->(EOF())
  lAddedAsSold:=.f.
  if !empty(NEWUPS->beback1)
    IF NEWUPS->BEBACK1 >=dBEGDATE
     IF NEWUPS->BEBACK1 <=dENDDATE
		  IF NEWUPS->UPTYPE=1
            nu3:=nu3+1
				AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

				IF NEWUPS->UPSTATUS='S'
						nS3:=nS3+1
						AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
						lAddedasSold:=.t.
				ENDIF
		  ENDIF
		  IF NEWUPS->UPTYPE=2
				IF NEWUPS->UPSTATUS='S'
						nS4:=nS4+1
						AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
						lAddedasSold:=.t.
				ENDIF
				nU4:=nU4+1
				AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)
        ENDIF
      ENDIF
    ENDIF
  endif
  if !empty(NEWUPS->beback2)
    IF NEWUPS->BEBACK2 >=dBEGDATE
     IF NEWUPS->BEBACK2 <=dENDDATE
			IF NEWUPS->UPTYPE=1
            nu3:=nu3+1
				AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)
				IF !lAddedAsSold
				 IF NEWUPS->UPSTATUS='S'
						nS3:=nS3+1
						AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
						lAddedasSold:=.t.
				 ENDIF
				ENDIF
			ELSE    // IF THIS IS A TYPE 2
				IF !lAddedAsSold
				 IF NEWUPS->UPSTATUS='S'
					nS4:=nS4+1
					AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
				 ENDIF
				ENDIF
				AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)
				nU4:=nu4+1
         ENDIF
      ENDIF
    ENDIF
  endif
  SKIP ALIAS NEWUPS
ENDDO

//DC_ARRAYVIEW(aSoldarray)

PRTSTATS('ALL UPTYPES FOR ' + DTOC(dBEGDATE)+' '+'TO ' +DTOC(dENDDATE),cFranchise,nU1,nU2,nU3,nU4,nU5,nU6,;
	nS1,nS2,nS3,nS4,nS5,nS6,nCP1,nCP2,nCP3,;
	nCP4,nCP5,nCP6,nNEW,nNEWS,nUSED,nUSEDS,nDDN,nDDU,nDemo,nTOmgr,nIupsIn,nPupsIn,aSoldArray,aUpsInarray)
CLOSE DATABASES
RETURN .T.

FUNCTION CHKSMI(cSM)

  IF EMPTY(cSM)
    RETURN .F.
  ENDIF
  SELECT SM
  seek cSM
  IF !found()
	RETURN FALSE
  ENDIF
  RETURN .T.

FUNCTION ICLOSEREPT()
LOCAL GETLIST:={} , getoptions, cFranchise ,xStatus , dBegdate, dEnddate , cSM ,oDialog
local nU1,nU2,nU3,nU4,nU5,nU6,nS1,nS2,nS3,nS4,nS5,nS6,nCP1,nCP2,nCP3,nCP4,nCP5,nCP6,nNEW,nNEWS,nUSED,nUSEDS,nDDN,nDDU,nDemo,nTOmgr,;
	nPupsIn,nIupsIn,aSoldArray:={}, lAddedAsSold:=.f. ,aUpsInarray:={}



nU1:=nU2:=nU3:=nU4:=nU5:=nU6:=nS1:=nS2:=nS3:=nS4:=nS5:=nS6:=nCP1:=nCP2:=nCP3:=nCP4:=nCP5:=nCP6:=nNEW;
	:=nNEWS:=nUSED:=nUSEDS:=nDDN:=nDDU:=nDemo:=nTOmgr:=nPupsIn:=nIupsIn:=0
cSM:=space(2)

IF !UseDb({'NEWUPS','SM'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF

NEWUPS->(ORDSETFOCUS('NEWUPSM'))

DCGETOPTIONS SAYWIDTH 0 AUTORESIZE sayfont '12.Courier New' getfont '12.Courier New'

dBEGDATE:=dENDDATE:=CTOD("  /  /  ")
@ 10,30 DCSAY 'Enter Beginning Date '  GET dBEGDATE valid{|| chkBEGDT(dBEGDATE)} SAYFONT '12.COURIER' GETFONT '12.COURIER'
@ 11,30 DCSAY 'Enter Ending Date    ' GET dENDDATE SAYFONT '12.COURIER'  GETFONT '12.COURIER'
@ 12,30 DCSAY 'Enter Salesperson    ' GET cSM VALID{|| CHKSMI(cSM)}
DCREAD GUI modal ENTEREXIT OPTIONS GETOPTIONS BUTTONS 2 TO XSTATUS FIT TITLE 'UPCOUNT Report'  eval{|o| setappwindow(o)}
IF !XSTATUS
  RETURN NIL
ENDIF
SELECT newups
GOTO TOP
ORDSETFOCUS('NEWUPSM')
DBGOTOP()
seek cSM
odialog:=DC_WAITON('Now Gathering Up Data')
DO WHILE !NEWUPS->(EOF())
  IF NEWUPS->SALESMAN=cSM
   IF NEWUPS->DATEIN=NEWUPS->BEBACK3 /// BEBACK3 IS REALLY THE ORIGINAL DATE IN ONLY COUNT AS A 1OR23OR4 IF THESE DATES ARE THE SAME
    IF NEWUPS->DATEIN >=dBEGDATE
     IF NEWUPS->DATEIN <=dENDDATE
		 IF NEWUPS->DATEIN <> NEWUPS->BEBACK1       /// DO NOT COUNT AN UP IF THE DATEIN AND BEBACK DATES ARE THE SAME
			IF NEWUPS->DATEIN <> NEWUPS->BEBACK2
		  		IF NEWUPS->UPTYPE = 1
        		        nU1:=nU1+1
						  AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 1'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)
		  				  IF NEWUPS->UPSTATUS='S'
		  						nS1:=nS1+1
								AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 1'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)

		  				  ENDIF
        		ENDIF
        		IF NEWUPS->UPTYPE = 2
        		        nU2:=nU2+1
						  AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 2'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

		  				  IF NEWUPS->UPSTATUS='S'
		  						nS2:=nS2+1
								AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 2'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
		  				  ENDIF

        		ENDIF
        		IF NEWUPS->UPTYPE = 3
        		        nU3:=nU3+1
						  AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

		  				  IF NEWUPS->UPSTATUS='S'
		  						nS3:=nS3+1
								AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
		  				  ENDIF
        		ENDIF
        		IF NEWUPS->UPTYPE = 4
        		        nU4:=nU4+1
						  AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

		  				  IF NEWUPS->UPSTATUS='S'
		  						nS4:=nS4+1
								AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
		  				  ENDIF
          	ENDIF
		   ENDIF     /// <> BEBACK1
		 ENDIF  ///     <> BEBACK2
		IF NEWUPS->ORIGSOURCE = 'I'
                nU5:=nU5+1
					 IF NEWUPS->UPSTATUS='S'
						nS5:=nS5+1
					 ENDIF
					 IF NEWUPS->UPTYPE < 5
						nIupsIn++
					 ENDIF
      ENDIF
		IF NEWUPS->ORIGSOURCE = 'P'
                nU6:=nU6+1
					 IF NEWUPS->UPSTATUS='S'
						nS6:=nS6+1
					 ENDIF
					 IF NEWUPS->UPTYPE < 5
						nPupsIn++
					 ENDIF

      ENDIF

      IF NEWUPS->NEWUSED='N'
          nNEW:=nNEW+1
          IF NEWUPS->upstatus='S'
            nNEWS:=nNEWS+1
          ENDIF
          IF NEWUPS->DOWNDEAL > 0
             nDDN:=nDDN+1
          ENDIF
      ENDIF
      IF NEWUPS->NEWUSED='U'
          nUSED:=nUSED+1
          IF NEWUPS->upstatus='S'
            nUSEDS:=nUSEDS+1
          ENDIF
          IF NEWUPS->DOWNDEAL > 0
             nDDU:=nDDU+1
          ENDIF
       ENDIF
		 if NEWUPS->demo='Y'
					nDemo++
		 endif
		 if !empty(NEWUPS->tomgr)
					nTOmgr++
		 endif

     ENDIF   //DATE
    ENDIF  //DATE
   ENDIF // BEBACK3=ORIGDATEIN
  ENDIF  //SM
  SKIP ALIAS NEWUPS
ENDDO
SELECT NEWUPS
GOTO BOTTOM
SKIP -1000
DO WHILE !NEWUPS->(EOF())
 lAddedAsSold:=.f.
 IF NEWUPS->SALESMAN=cSM
  if !empty(NEWUPS->beback1)
    IF NEWUPS->BEBACK1 >=dBEGDATE
     IF NEWUPS->BEBACK1 <=dENDDATE
		IF NEWUPS->UPTYPE=1
				IF NEWUPS->UPSTATUS='S'
						nS3:=nS3+1
						AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
						lAddedAsSold:=.t.
				ENDIF
            nu3:=nu3+1
				AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

		ENDIF
		IF NEWUPS->UPTYPE=2
				IF NEWUPS->UPSTATUS='S'
						IF !lAddedAsSold
						 nS4:=nS4+1
						 AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
						ENDIF
				ENDIF
				nU4:=nu4+1
				AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

      ENDIF
	  ENDIF
    ENDIF
  endif
  if !empty(NEWUPS->beback2)
    IF NEWUPS->BEBACK2 >=dBEGDATE
     IF NEWUPS->BEBACK2 <=dENDDATE
		IF NEWUPS->UPTYPE=1
            nu3:=nu3+1
				AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)

				IF !lAddedAsSold
				 IF NEWUPS->UPSTATUS='S'
					nS3:=nS3+1
					AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 3'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
				 ENDIF
				ENDIF
			ELSE
				IF !lAddedAsSold
				 IF NEWUPS->UPSTATUS='S'
					nS4:=nS4+1
					AADD(aSoldarray,dtoc(NEWUPS->datein)+'  '+substr(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->STOCKNUM+' '+NEWUPS->NEWUSED)
				 ENDIF
				ENDIF
				nU4:=nu4+1
				AADD(aUpsInarray,dtoc(NEWUPS->datein)+'  '+SUBSTR(NEWUPS->LASTNAME,1,15)+' 4'+' '+NEWUPS->ORIGSOURCE+' '+NEWUPS->INTEREST+' '+NEWUPS->NEWUSED)
      ENDIF
     ENDIF
    ENDIF                //bb2
  endif        //// emoty
 endif  /// smi
  SKIP ALIAS NEWUPS
ENDDO
DC_IMPL(oDialog)

PRTSTATS('ALL UPTYPES FOR ' + DTOC(dBEGDATE)+' '+'TO ' +DTOC(dENDDATE),cSM,nU1,nU2,nU3,nU4,nU5,nU6,;
	nS1,nS2,nS3,nS4,nS5,nS6,nCP1,nCP2,nCP3,;
	nCP4,nCP5,nCP6,nNEW,nNEWS,nUSED,nUSEDS,nDDN,nDDU,nDemo,nTOmgr,nIupsIn,nPupsIn,aSoldArray,aUpsInArray)
CLOSE DATABASES
RETURN .T.




FUNCTION PRTSTATS(MESS,cFranchise,nU1,nU2,nU3,nU4,nU5,nU6,nS1,nS2,nS3,nS4,nS5,nS6,nCP1,nCP2,nCP3,nCP4,nCP5,nCP6,;
		nNEW,nNEWS,nUSED,nUSEDS,nDDN,nDDU,nDemo,nTOmgr,nIupsIn,nPupsIn,aSoldarray,aUpsInarray)
LOCAL GETLIST:={} , i  , nNumSld:=0  ,nNumIn:=0
local nCopies:=1,nOrient:=1,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New'
TOP_MAR(2)
BOT_MAR(2)
PAGE_LEN(62)

IF !Print_Choice('Up Report', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
ENDIF
IF !PrintOn('Up Report', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
ENDIF



@ DC_PRINTERROW()+1,0 DCPRINT SAY MESS    FONT '14.COURIER'
IF !EMPTY(cFranchise)
  @ DC_PRINTERROW()+1,0 DCPRINT SAY cFranchise FONT '12.COURIER BOLD'
ENDIF
IF EMPTY(cFranchise)
  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'ALL FRANCHISES' FONT '12.COURIER BOLD'
ENDIF

@ DC_PRINTERROW()+2,0 DCPRINT SAY 'Percentages By Up Type'
@ DC_PRINTERROW()+2,0 DCPRINT SAY 'UPTYPE 1s' + STR(nU1)
@ DC_PRINTERROW(),DC_PRINTERCOL()+10 DCPRINT SAY 'SOLD 1s' + STR(nS1)
@ DC_PRINTERROW(),DC_PRINTERCOL()+10 DCPRINT SAY 'CLOSING% 1s'
@ DC_PRINTERROW(),DC_PRINTERCOL()+5 DCPRINT SAY (nS1/nU1) * 100 PICTURE '999.99' FONT '10.COURIER BOLD'
@ DC_PRINTERROW(),DC_PRINTERCOL()+10 DCPRINT SAY 'BEBACK% 1s '
@ DC_PRINTERROW(),DC_PRINTERCOL()+5 DCPRINT SAY (nU3/(nU1-nS1)) *100 PICTURE '999.99' FONT '10.COURIER BOLD'


@ DC_PRINTERROW()+2,0 DCPRINT SAY 'UPTYPE 2s' + STR(nU2)
@ DC_PRINTERROW(),DC_PRINTERCOL()+10 DCPRINT SAY 'SOLD 2s' + STR(nS2)
@ DC_PRINTERROW(),DC_PRINTERCOL()+10 DCPRINT SAY 'CLOSING% 2s'
@ DC_PRINTERROW(),DC_PRINTERCOL()+5 DCPRINT SAY (nS2/nU2) * 100 PICTURE '999.99' FONT '10.COURIER BOLD'
@ DC_PRINTERROW(),DC_PRINTERCOL()+10 DCPRINT SAY 'BEBACK% 2s '
@ DC_PRINTERROW(),DC_PRINTERCOL()+5 DCPRINT SAY (nU4/(nU2-nS2)) *100 PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+2,0 DCPRINT SAY 'UPTYPE 3s' + STR(nU3)
@ DC_PRINTERROW(),DC_PRINTERCOL()+10 DCPRINT SAY 'SOLD 3s' + STR(nS3)
@ DC_PRINTERROW(),DC_PRINTERCOL()+10 DCPRINT SAY 'CLOSING% 3s'
@ DC_PRINTERROW(),DC_PRINTERCOL()+5 DCPRINT SAY (nS3/nU3) * 100 PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+2,0 DCPRINT SAY 'UPTYPE 4s' + STR(nU4)
@ DC_PRINTERROW(),DC_PRINTERCOL()+10 DCPRINT SAY 'SOLD 4s' + STR(nS4)
@ DC_PRINTERROW(),DC_PRINTERCOL()+10 DCPRINT SAY 'CLOSING% 4s'
@ DC_PRINTERROW(),DC_PRINTERCOL()+5 DCPRINT SAY (nS4/nU4) * 100 PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+3,0 DCPRINT SAY 'Percentages By Up Source'

@ DC_PRINTERROW()+2,0 DCPRINT SAY 'Floor Traffic Recorded' + alltrim(STR(nU1+nU2+nU3+nU4-(nIupsIn+nPupsIn)))
@ DC_PRINTERROW(),60 DCPRINT SAY 'SOLD ' + STR((nS1+nS2+nS3+nS4)-(nS5+nS6))
@ DC_PRINTERROW(),90 DCPRINT SAY 'CLOSING% '
@ DC_PRINTERROW(),100 DCPRINT SAY ((nS1+nS2+nS3+nS4)-(nS5+nS6))/((nU1+nU2+nU3+nU4)-(nIupsIn+nPupsIn)) * 100 PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+2,0 DCPRINT SAY 'Internet Ups Recorded ' + alltrim(STR(nU5))
@ DC_PRINTERROW(),30 DCPRINT SAY 'Converted  '+ alltrim(str(nIupsIn))
@ DC_PRINTERROW(),60 DCPRINT SAY 'SOLD ' + STR(nS5)
@ DC_PRINTERROW(),90 DCPRINT SAY 'CLOSING% '
@ DC_PRINTERROW(),100 DCPRINT SAY (nS5/nIupsIn) * 100 PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+2,0 DCPRINT SAY 'Phone Ups Recorded    ' + alltrim(STR(nU6))
@ DC_PRINTERROW(),30 DCPRINT SAY 'Converted  '+ alltrim(str(nPupsIn))
@ DC_PRINTERROW(),60 DCPRINT SAY 'SOLD ' + STR(nS6)
@ DC_PRINTERROW(),90 DCPRINT SAY 'CLOSING% '
@ DC_PRINTERROW(),100 DCPRINT SAY (nS6/nPupsIn) * 100 PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+2,0 DCPRINT SAY 'Demo Rides Recorded '
@ DC_PRINTERROW(),60 DCPRINT SAY nDemo picture '9999' FONT '12.COURIER BOLD'


@ DC_PRINTERROW()+2,0 DCPRINT SAY 'TO s Recorded '
@ DC_PRINTERROW(),60 DCPRINT SAY nTOmgr picture '9999' FONT '12.COURIER BOLD'

skipaline()

@ DC_PRINTERROW()+1,0 DCPRINT SAY 'TOTAL UPS FOR PERIOD '
@ DC_PRINTERROW(),60 DCPRINT SAY (nU1+nU2+nU3+nU4) PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+1,0 DCPRINT SAY 'TOTAL SOLD FOR PERIOD '
@ DC_PRINTERROW(),60 DCPRINT SAY (nS1+nS2+nS3+nS4) PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+1,0 DCPRINT SAY 'TOTAL UPS CLOSING RATIO '
@ DC_PRINTERROW(),60 DCPRINT SAY (nS1+nS2+nS3+nS4)/(nU1+nU2+nU3+nU4) * 100 PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+1,0 DCPRINT SAY 'TOTAL NEW VEHICLE UPS FOR PERIOD '
@ DC_PRINTERROW(),60 DCPRINT SAY (nNEW) PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+1,0 DCPRINT SAY 'TOTAL USED VEHICLE UPS FOR PERIOD '
@ DC_PRINTERROW(),60 DCPRINT SAY (nUSED) PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+1,0 DCPRINT SAY 'TOTAL NEW VEHICLE UPS SOLD FOR PERIOD '
@ DC_PRINTERROW(),60 DCPRINT SAY (nNEWS) PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+1,0 DCPRINT SAY 'TOTAL USED VEHICLE UPS SOLD FOR PERIOD '
@ DC_PRINTERROW(),60 DCPRINT SAY (nUSEDS) PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+1,0 DCPRINT SAY 'TOTAL NEW VEHICLE DEALS DOWN FOR PERIOD '
@ DC_PRINTERROW(),60 DCPRINT SAY (nDDN) PICTURE '999.99' FONT '10.COURIER BOLD'

@ DC_PRINTERROW()+1,0 DCPRINT SAY 'TOTAL USED VEHICLE DEALS DOWN FOR PERIOD '
@ DC_PRINTERROW(),60 DCPRINT SAY (nDDU) PICTURE '999.99' FONT '10.COURIER BOLD'

skipaline()
skipaline()
nNumSld:=0
@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Date    Sold Name        Type  Source Stk#       N/U'
FOR i := 1 TO len(aSoldArray)
	nNumSld++
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 	alltrim(str(nNumsld))+' '+aSoldarray[i]
	IF DCPAGEEJECT()
		@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Sold Ups'
	ENDIF
NEXT
skipaline()
skipaline()
nNumIn:=0
@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Date    Up Name        Type  Source Interest       N/U'
FOR i := 1 TO len(aUpsInArray)
	nNumIn++
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 	alltrim(str(nNumIn))+' '+aUpsInarray[i]
	IF DCPAGEEJECT()
		@ DC_PRINTERROW()+1,0 DCPRINT SAY 'Ups in Store'
	ENDIF
NEXT
PRINTOFF(oPrinter)
RETURN NIL


FUNCTION UPINDX()
	local oDialog , aStruct:={}
  oDIALOG:=DC_WAITON('NOW REINDEXING')

  IF !UseDb({'NEWUPS','HNEWUPS','ZIPPER','CLASSORT','ADSOURCE','SM','UCAPPRSL','INETUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
ENDIF



  DC_IMPL(oDIALOG)
  CLOSE DATABASES
RETURN NIL

FUNCTION AVVCOMPARE()
	LOCAL GETLIST:={},GETOPTIONS , nTotcount:=0, nTotinUps:=0 ,nTotPoss:=0 ,nTotsold:=0    , nRec:=0,cName,cEmail
	local nCopies:=1,nOrient:=2,nDefmode:=4, oPrinter,cOutfile:='', cFont:='8.Courier New' , nPage:=0
	TOP_MAR(2)
	BOT_MAR(2)
	PAGE_LEN(62)


 IF !UseDb({'NEWUPS','INETUPS',''})
 	 DC_WINALERT('Cannot Open Files     .')
    	 RETURN NIL
 ENDIF


	IF !Print_Choice('AVV COMPARE', @nCopies,' ',.f.,@nOrient,@cFont,@nDefMode,cOutfile)
	  RETURN NIL
   ENDIF
   IF !PrintOn('AVVCOMPARE', oPrinter, nOrient, cFont, nCopies,.t.)
	  RETURN NIL
   ENDIF

	_AVVSYNCHDR1(@nPage)

	SELECT INETUPS
	DO WHILE !INETUPS->(EOF())
		  nTotcount++
		  cEmail:=UPPER(INETUPS->EMAIL)
		  cName:=UPPER(INETUPS->LASTNAME)
		  SELECT NEWUPS
		  ORDSETFOCUS('NEWUPEML')
		  IF !EMPTY(cEmail)
		   SEEK cEmail
		   IF FOUND()
			 nTotinUps++
			 IF NEWUPS->UPSTATUS='S'
				nTotsold++
			 ENDIF
			 IF !NEWUPS->ORIGSOURCE='I'
				IF NEWUPS->(DC_RECLOCK())
					nRec:=NEWUPS->(RECNO())
					REPLACE NEWUPS->ORIGSOURCE WITH 'I'
					DBRUNLOCK(nRec)
				ENDIF
			 ENDIF
			 @ DC_PRINTERROW()+1,0 DCPRINT SAY SUBSTR(NEWUPS->LASTNAME,1,20)
			 @ DC_PRINTERROW(),25 DCPRINT SAY SUBSTR(NEWUPS->FIRSTNAME,1,10)
			 @ DC_PRINTERROW(),40 DCPRINT SAY NEWUPS->SALESMAN
			 @ DC_PRINTERROW(),45 DCPRINT SAY NEWUPS->UPTYPE
			 @ DC_PRINTERROW(),50 DCPRINT SAY NEWUPS->UPID
			 @ DC_PRINTERROW(),60 DCPRINT SAY SUBSTR(NEWUPS->EMAIL,1,30)
			 @ DC_PRINTERROW(),95 DCPRINT SAY NEWUPS->ORIGSOURCE
			 @ DC_PRINTERROW(),100 DCPRINT SAY NEWUPS->UPSTATUS
			 @ DC_PRINTERROW(),105 DCPRINT SAY INETUPS->CONLAST
		   ENDIF
		   IF DCPAGEEJECT()
			 _AVVSYNCHDR1(@nPage)
		   ENDIF
		  ENDIF
		  SKIP ALIAS INETUPS
	ENDDO
	DCPRINT EJECT PRINTER oPrinter

	_AVVSYNCHDR2(@nPage)

	SELECT INETUPS
	INETUPS->(DBGOTOP())
	DO WHILE !INETUPS->(EOF())
		  cName:=UPPER(INETUPS->LASTNAME)
		  SELECT NEWUPS
		  ORDSETFOCUS('NEWUPLNM')
		  SEEK cName
		  IF FOUND()
				IF ALLTRIM(UPPER(INETUPS->EMAIL)) <> ALLTRIM(UPPER(NEWUPS->EMAIL))
					nTotPoss++
					IF NEWUPS->UPSTATUS='S'
					 nTotsold++
			   	ENDIF
			   	IF !NEWUPS->ORIGSOURCE='I'
					 IF NEWUPS->(DC_RECLOCK())
						nRec:=NEWUPS->(RECNO())
						REPLACE NEWUPS->ORIGSOURCE WITH 'I'
						DBRUNLOCK(nRec)
					 ENDIF
			   	ENDIF

					@ DC_PRINTERROW()+1,0 DCPRINT SAY '**'
					@ DC_PRINTERROW(),3 DCPRINT SAY SUBSTR(NEWUPS->LASTNAME,1,20) FONT '8.Courier New Bold'
			   	@ DC_PRINTERROW(),25 DCPRINT SAY SUBSTR(NEWUPS->FIRSTNAME,1,10)
			   	@ DC_PRINTERROW(),40 DCPRINT SAY NEWUPS->SALESMAN
			   	@ DC_PRINTERROW(),45 DCPRINT SAY NEWUPS->UPTYPE
			   	@ DC_PRINTERROW(),50 DCPRINT SAY NEWUPS->UPID
			   	@ DC_PRINTERROW(),60 DCPRINT SAY SUBSTR(NEWUPS->EMAIL,1,30)
			   	@ DC_PRINTERROW(),95 DCPRINT SAY NEWUPS->ORIGSOURCE
			   	@ DC_PRINTERROW(),100 DCPRINT SAY NEWUPS->UPSTATUS
					@ DC_PRINTERROW(),105 DCPRINT SAY ALLTRIM(INETUPS->FIRSTNAME)+' '+ALLTRIM(INETUPS->LASTNAME)
					IF DCPAGEEJECT()
			   	 _AVVSYNCHDR2(@nPage)
		      	ENDIF
				ENDIF  /// EMAILS <>
		  ENDIF     /// FOUND
		  SELECT INETUPS
		  SKIP ALIAS INETUPS
	ENDDO
	DCPRINT EJECT PRINTER oPrinter


	SELECT INETUPS
	INETUPS->(DBGOTOP())
	_AVVSYNCHDR3(@nPage)

	DO WHILE !INETUPS->(EOF())
		  cName:=UPPER(INETUPS->LASTNAME)
		  SELECT NEWUPS
		  ORDSETFOCUS('NEWUPLNM')
		  SEEK cName
		  IF !FOUND()

				@ DC_PRINTERROW()+1,0 DCPRINT SAY ALLTRIM(INETUPS->FIRSTNAME)+' '+ALLTRIM(INETUPS->LASTNAME)+','+ALLTRIM(INETUPS->CITY)
				@ DC_PRINTERROW(),40 DCPRINT SAY ALLTRIM(INETUPS->EMAIL)
				@ DC_PRINTERROW(),80 DCPRINT SAY INETUPS->TIMESTAMP
				@ DC_PRINTERROW(),90 DCPRINT SAY INETUPS->SOURCE
				@ DC_PRINTERROW(),120 DCPRINT SAY INETUPS->CONLAST
		  ENDIF
		  SELECT INETUPS
		  INETUPS->(DBSKIP())
		  IF DCPAGEEJECT()
			_AVVSYNCHDR3(@nPage)
		  ENDIF
	  ENDDO
	  SKIPALINE()
	  IF DCPAGEEJECT()
		 _AVVSYNCHDR2(@nPage)
	  ENDIF
	  SKIPALINE()
	  IF DCPAGEEJECT()
		 _AVVSYNCHDR2(@nPage)
	  ENDIF
	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Leads Matched by eMail '+ str(nTotinups)
	  IF DCPAGEEJECT()
		 _AVVSYNCHDR2(@nPage)
	  ENDIF
	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Leads Matched by Lastname '+ str(nTotposs)
	  IF DCPAGEEJECT()
		 _AVVSYNCHDR2(@nPage)
	  ENDIF
	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Internet Leads Sold '+ str(nTotsold)
	  IF DCPAGEEJECT()
		 _AVVSYNCHDR2(@nPage)
	  ENDIF


	  @ DC_PRINTERROW()+1,0 DCPRINT SAY 'Total Leads Processed '+ str(nTotcount)

	  PRINTOFF(oPrinter)
	  return nil


STATIC FUNCTION _AVVSYNCHDR1(nPage)
	nPage++
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'AVV LEAD SOURCE WITH COMMON EMAIL ADDRESS ' + alltrim(str(nPage))
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'LASTNAME'
	@ DC_PRINTERROW(),25 DCPRINT SAY 'FIRSTNAME'
	@ DC_PRINTERROW(),40 DCPRINT SAY 'SM'
	@ DC_PRINTERROW(),45 DCPRINT SAY 'TYPE'
	@ DC_PRINTERROW(),50 DCPRINT SAY 'UPID'
	@ DC_PRINTERROW(),60 DCPRINT SAY 'EMAIL'
	@ DC_PRINTERROW(),95 DCPRINT SAY 'SRCE'
	@ DC_PRINTERROW(),100 DCPRINT SAY 'STAT'
RETURN NIL

STATIC FUNCTION _AVVSYNCHDR2(nPage)
	nPage++
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'AVV LEAD SOURCE WITH COMMON LASTNAME** ' + alltrim(str(nPage))
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'LASTNAME'
	@ DC_PRINTERROW(),25 DCPRINT SAY 'FIRSTNAME'
	@ DC_PRINTERROW(),40 DCPRINT SAY 'SM'
	@ DC_PRINTERROW(),45 DCPRINT SAY 'TYPE'
	@ DC_PRINTERROW(),50 DCPRINT SAY 'UPID'
	@ DC_PRINTERROW(),60 DCPRINT SAY 'EMAIL'
	@ DC_PRINTERROW(),95 DCPRINT SAY 'SRCE'
	@ DC_PRINTERROW(),100 DCPRINT SAY 'STAT'
	RETURN NIL

STATIC FUNCTION _AVVSYNCHDR3(nPage)
	nPage++
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'AVV LEAD SOURCE WITH NO MATCHING RECORDS IN UPS ' + alltrim(str(nPage))
	@ DC_PRINTERROW()+1,0 DCPRINT SAY 'LEAD NAME / CITY'
	@ DC_PRINTERROW(),40 DCPRINT SAY 'LEAD EMAIL'
	@ DC_PRINTERROW(),80 DCPRINT SAY 'LEAD DATE'
	@ DC_PRINTERROW(),120 DCPRINT SAY 'SALESPERSON'
	RETURN NIL





FUNCTION BROWBYALL()
  LOCAL GETLIST:={}  , oBrowse, getoptions ,xstatus, aPres , cSearchkey:=space(15), cIndexord:='Prospect Last Name'
  LOCAL cLastName:=SPACE(25)    , cProgram, oToolbar,oToolbar2 ,oRecord
  cProgram:='BROWBYNM'
  IF.NOT.SECURITY(@cProgram)
    DC_winALERT('SECURITY VIOLATION')
    RETURN NIL
  ENDIF




  IF !OPENUPSFILES()
	CLOSE ALL
   RETURN NIL
  ENDIF



  NEWUPS->(ORDSETFOCUS('NEWUPNM'))
  NEWUPS->(DBGOTOP())

  DCGETOPTIONS  ;
  SAYWIDTH 0 ;
  AUTORESIZE SAYFONT '10.Courier New' getfont '10.Courier New'

aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE }, ;//     Header FG Color       ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY },; //  Header BG Color       ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },;   //Row Sep       ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED }, ;  //Col Sep       ;
    { XBP_PP_COL_DA_ROWHEIGHT, 45 },                ; //  Row Height      ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, BD_PALEBLUE },;  // HILITE BG Color   ;
	 { XBP_PP_COL_DA_HILITE_FGCLR, GRA_CLR_BLACK},;  // HILITE BG Color   ;
    { XBP_PP_COL_DA_CELLHEIGHT, 10 }  }              // Cell Height


 @ 1,0 DCSAY 'Current Order' get cIndexord editprotect{|| TRUE }


 @ 1,50 DCSAY 'Enter Search Key For Order Chosen' get cSearchkey GETID 'SEARCH' PICTURE '@!' ;
		      KEYBLOCK {|a,b,o| DBSELECTAREA('NEWUPS'), DC_BrowseAutoSeek(a,o,oBrowse)}


 @ 3,0 DCTOOLBAR oToolbar SIZE 132,1.5   BUTTONSIZE 18,1.5
 @ 5,0 DCTOOLBAR oToolbar2 SIZE 132, 1.5 BUTTONSIZE 18,1.5
 @ 7,0 DCSAY 'Any column header that starts with (**) can be selected for sort order. RIGHT click the desired column header to sort on.'

@ 8,0 DCSAY 'COLOR CODES'
@ 8,15 DCSAY 'Sold' SAYCOLOR 1,BD_KEYLIME
@ 8,30 DCSAY 'Active' SAYCOLOR 1, BD_WASHGREY
@ 8,45 DCSAY 'Tickle' SAYCOLOR 1,BD_CORNFLOWERBLUE
@ 8,60 DCSAY 'DEAD' SAYCOLOR 1,BD_SALMON
@ 8,75 DCSAY 'Rescue' SAYCOLOR 2,BD_SUNNYYELLOW
@ 8,90 DCSAY 'PhoneUp' SAYCOLOR 2,BD_FADEDPURPLE
@ 8,105 DCSAY 'InterNetUp' SAYCOLOR 1,BD_BLUEMIST


  DCADDBUTTON CAPTION 'Search By Date'                              ;
		ACTION {|| _nlook4date(), oBrowse:REFRESHALL(),obrowse:forcestable()}        ;
		PARENT oToolBar                                     ;
		TOOLTIP 'Search Ups By Date'

  DCADDBUTTON CAPTION 'Print UPSheet '                              ;
	ACTION {||PRINTUPSHEET(NEWUPS->(RECNO()))};
	PARENT oToolBar

  DCADDBUTTON CAPTION 'Send eMail '                              ;
	ACTION {||upemail()};
	PARENT oToolBar

  DCADDBUTTON CAPTION 'Review eMails '                              ;
		ACTION {|| BROWEMAILLOG('',NEWUPS->EMAIL)} ;
		PARENT oToolBar2




DCADDBUTTON CAPTION 'Desking Tool;Quote '                              ;
	ACTION  {||	UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)};
	PARENT oToolBar2

DCADDBUTTON CAPTION 'New Vehicle;Browser'                              ;
	ACTION {||SBROWBYDESC(),DC_GETREFRESH(GETLIST)  };
	PARENT oToolBar2

DCADDBUTTON CAPTION 'Used Vehicle;Browser'                              ;
	ACTION {||SBROWBYDESC(),DC_GETREFRESH(GETLIST)  };
	PARENT oToolBar2



DCADDBUTTON CAPTION 'Comments '                              ;
	ACTION {|| UPNOTES(),DC_GETREFRESH(GETLIST),SETAPPFOCUS(obROWSE:GETCOLUMN(1))  };
	PARENT oToolBar2

/* ----- Create browse ----- */

@ 10,1 DCBROWSE oBrowse ALIAS 'NEWUPS'                 ;
	SIZE 170,20                                     ;
	PRESENTATION aPres;
	THUMBLOCK 500 ;
	FONT '8.Lucida Console'  ;
	HEADLINES 4 ;
	;//EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITACROSS ;
	DATALINK{|| UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)};
	sortscolor 0,2;
	SORTUCOLOR 1,6

DCBROWSECOL FIELD NEWUPS->SALESMAN ;
	HEADER "**SALESMAN" PARENT oBrowse  ;
	WIDTH 3      ;
	SORT{|| cIndexord:='Sales Person Id   ',NEWUPS->(ORDSETFOCUS('NEWUPSM')),NEWUPS->(DBGOTOP()), OBROWSE:REFRESHALL(),oBrowse:forcestable(),DC_GETREFRESH(GETLIST)}

DCBROWSECOL FIELD NEWUPS->DATEIN ;
	HEADER "**DATEIN" PARENT oBrowse  ;
	WIDTH 8     ;
	SORT{|| cIndexord:='Up Date in YYYYMMDD ',NEWUPS->(ORDSETFOCUS('NEWUPDT')),NEWUPS->(DBGOTOP()),OBROWSE:REFRESHALL(),oBrowse:forcestable(),DC_GETREFRESH(GETLIST)}


DCBROWSECOL FIELD NEWUPS->UPSTATUS ;
	HEADER "STATUS" PARENT oBrowse COLOR {|| getupcolor(NEWUPS->UPSTATUS)} ;
	WIDTH 1

DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "**UPID/PHN#" PARENT oBrowse ; //COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 7 ;
  	SORT{|| cIndexord:='Up Id Index       ',NEWUPS->(ORDSETFOCUS('NEWUPIN')),NEWUPS->(DBGOTOP()), OBROWSE:REFRESHALL(),oBrowse:forcestable(),DC_GETREFRESH(GETLIST)}


DCBROWSECOL DATA {{|| NEWUPS->LASTNAME},;
		               {|| NEWUPS->FIRSTNAME},;
							{|| NEWUPS->STREET },;
							{|| NEWUPS->CITYSTATE+' '+NEWUPS->ZIPCODE}};
	WIDTH 25 HEADER "LastName;FirstName;Street;CitySt" PARENT oBrowse ;
	EDITPROTECT{|| TRUE }   ;
	SORT{|| cIndexord:='Prospect Last Name   ',NEWUPS->(ORDSETFOCUS('NEWUPNM')),NEWUPS->(DBGOTOP()), OBROWSE:REFRESHALL(),oBrowse:forcestable(),DC_GETREFRESH(GETLIST)}


DCBROWSECOL DATA {{|| TRANSFORM(NEWUPS->AREA+NEWUPS->UPID,'@R 999.999.9999')},;
						{|| TRANSFORM(NEWUPS->WORKPHN,'@R 999.999.9999.99999')},;
						{|| TRANSFORM(NEWUPS->CELLPHONE,'@R 999.999.9999')},;
						{|| NEWUPS->email}};
	HEADER 'Phone;WorkPhn;CellPhn;eMail';
	width 25 PARENT oBrowse ;
	EDITPROTECT{|| TRUE }




DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "**TYPE" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1;
	SORT{|| cIndexord:='Up Type            ',NEWUPS->(ORDSETFOCUS('NEWUPTYP')),NEWUPS->(DBGOTOP()), OBROWSE:REFRESHALL(),oBrowse:forcestable(),DC_GETREFRESH(GETLIST)}

DCBROWSECOL FIELD NEWUPS->ORIGSOURCE ;
	HEADER "**Source" PARENT oBrowse COLOR {|| getphoneupcolor(NEWUPS->ORIGSOURCE)} ;
	WIDTH 1 ;
	SORT{|| cIndexord:='Up Source          ',NEWUPS->(ORDSETFOCUS('NEWUPSRC')),NEWUPS->(DBGOTOP()), OBROWSE:REFRESHALL(),oBrowse:forcestable(),DC_GETREFRESH(GETLIST)}




DCBROWSECOL DATA { ;
                 {|| NEWUPS->BEBACK1},;
	              {|| NEWUPS->BEBACK2}};
	HEADER "BeBack1;BeBack2" PARENT oBrowse  ;
	WIDTH 8

DCBROWSECOL DATA { ;
						{|| IIF(NEWUPS->NEWUSED='N' ,'NEW' ,'USED' )},;
	               {|| NEWUPS->INTEREST}};
	HEADER "New/Used;Interest" PARENT oBrowse  ;
	WIDTH 10 ;
	SORT{|| cIndexord:='Up Interest Order ',NEWUPS->(ORDSETFOCUS('NEWUPINT')),NEWUPS->(DBGOTOP()), OBROWSE:REFRESHALL(),oBrowse:forcestable(),DC_GETREFRESH(GETLIST)}

DCBROWSECOL DATA { ;
						{|| NEWUPS->ADSOURCE},;
	               {|| NEWUPS->TOMGR}};
	HEADER "Source/TOMgr;Interest" PARENT oBrowse  ;
	WIDTH 8


DCBROWSECOL DATA { ;
						{|| NEWUPS->COMMENTS},;
	               {|| NEWUPS->COMMENT2}};
	HEADER "Comments" PARENT oBrowse  ;
	WIDTH 20

 DCBROWSECOL DATA { ;
						{|| NEWUPS->SALECODE},;
	               {|| NEWUPS->DOWNDEAL}};
	HEADER "New/Used;Interest" PARENT oBrowse  ;
	WIDTH 1


DCREAD GUI ;
	FIT ;
	BUTTONS DCGUI_BUTTON_EXIT  ;
	OPTIONS GETOPTIONS  ;
	EVAL{|| SETAPPFOCUS(DC_GETOBJECT(GETLIST,'SEARCH'))};
	TITLE "Up Browser"
CLOSE ALL
ReTURN nil

static function _nLook4name()
	local getlist:={}, getoptions , cName:=space(10)
	DCGEToptions TABSTOP saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	@ 10,0 DCSAY 'Enter Name to Search For.' get cName picture '@!'
	DCREAD GUI MODAL ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Name Search' EVAL{|o| SETAPPWINDOW(o)}
	SELECT NEWUPS
	NEWUPS->(ORDSETFOCUS('NEWUPNM'))
	SET SOFTSEEK ON
	NEWUPS->(DBGOTOP())
	SEEK cName
	SET SOFTSEEK OFF
RETURN NIL

static function _nLook4upid()
	local getlist:={}, getoptions , cUpId:=space(7)
	DCGEToptions TABSTOP saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	@ 10,0 DCSAY 'Enter Phone to Search For.' get cUpid
	DCREAD GUI MODAL ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Name Search' EVAL{|o| SETAPPWINDOW(o)}
	SELECT NEWUPS
	NEWUPS->(ORDSETFOCUS('NEWUPIN'))
	SET SOFTSEEK ON
	NEWUPS->(DBGOTOP())
	SEEK cUpid
	SET SOFTSEEK OFF
RETURN NIL

static function _nLook4date()
	local getlist:={}, getoptions , dDate:=ctod('  /  /  ')
	DCGEToptions TABSTOP saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	@ 10,0 DCSAY 'Enter Date to Search For.' get dDate
	DCREAD GUI MODAL ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Name Search' EVAL{|o| SETAPPWINDOW(o)}
	SELECT NEWUPS
	NEWUPS->(ORDSETFOCUS('NEWUPDT'))
	SET SOFTSEEK ON
	NEWUPS->(DBGOTOP())
	SEEK dDate
RETURN NIL

static function _nLook4salesman()
	local getlist:={}, getoptions , cSM:=space(3)
	DCGEToptions TABSTOP saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	@ 10,0 DCSAY 'Enter SalesPerson to Search For.' get cSM picture '@!'
	DCREAD GUI MODAL ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Name Search' EVAL{|o| SETAPPWINDOW(o)}
	SELECT NEWUPS
	NEWUPS->(ORDSETFOCUS('NEWUPSM'))
	SET SOFTSEEK ON
	NEWUPS->(DBGOTOP())
	SEEK cSm
	SET SOFTSEEK OFF
	RETURN NIL


FUNCTION HBROWBYNM()
  LOCAL GETLIST:={}  , oBrowse, cUpid:=space(7),cName:=space(10),dDate:=ctod('  /  /  ') ,cSM:=space(3), getoptions ,xstatus, aPres
  LOCAL cLastName:=SPACE(25)    , cProgram ,oToolbar
  cProgram:='BROWBYNM'
  IF.NOT.SECURITY(@cProgram)
    DC_winALERT('SECURITY VIOLATION')
    RETURN NIL
  ENDIF
  IF !UseDb({'HNEWUPS'})
  	 DC_WINALERT('Cannot Open Files     .')
     	 RETURN NIL
  ENDIF
  DCGETOPTIONS  ;
  SAYWIDTH 0 ;
  AUTORESIZE SAYFONT '10.Courier New' getfont '10.Courier New' TABSTOP
 /*
@ 10,0 DCSAY 'ENTER THE LASTNAME YOU WISH TO BROWSE' GET cLastName SAYCOLOR 3,0 SAYFONT '10.COURIER' GETFONT '12.COURIER'
DCREAD GUI TO XSTATUS ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'ENTER LASTNAME TO BEGIN BROWSE'
IF !XSTATUS
  RETURN NIL
ENDIF
SELECT HNEWUPS
SET SOFTSEEK ON
SEEK UPPER(cLastName)
SET SOFTSEEK OFF  */

aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
    { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
	 { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

 @ 2,0 DCTOOLBAR oToolbar SIZE 100,1.5 BUTTONSIZE 20,1
 @ 5,0 DCSAY 'Any fields that start with an asterisk (*) can now be sorted. RIGHT click the header column to sort.'

 DCADDBUTTON CAPTION 'Search Name'                              ;
		SIZE 20                                              ;
		ACTION {|| _look4name(), oBrowse:REFRESHALL(),obrowse:forcestable()}        ;
		PARENT oToolBar                                     ;
		TOOLTIP 'Search Ups By name'

 DCADDBUTTON CAPTION 'Search PhoneID'                              ;
		SIZE 20                                              ;
		ACTION {|| _look4upid(), oBrowse:REFRESHALL(),obrowse:forcestable()}        ;
		PARENT oToolBar                                     ;
		TOOLTIP 'Search Ups By Phone'

 DCADDBUTTON CAPTION 'Search SalesPerson'                              ;
		SIZE 20                                              ;
		ACTION {|| _look4salesman(), oBrowse:REFRESHALL(),obrowse:forcestable()}        ;
		PARENT oToolBar                                     ;
		TOOLTIP 'Search Ups By Salesperson'

  DCADDBUTTON CAPTION 'Search By Date'                              ;
		SIZE 20                                              ;
		ACTION {|| _look4date(), oBrowse:REFRESHALL(),obrowse:forcestable()}        ;
		PARENT oToolBar                                     ;
		TOOLTIP 'Search Ups By Date'

  DCADDBUTTON CAPTION 'View Detail'                              ;
		SIZE 20                                              ;
		ACTION {|| HUPLOOK(HNEWUPS->(RECNO())), oBrowse:REFRESHALL(),obrowse:forcestable()}        ;
		PARENT oToolBar                                     ;
		TOOLTIP 'View Up Record'


/* give user chance to seek on vars
@ 2,1 DCSAY 'Lastname' get cName VALID{|| _LOOK4NAME(cName),oBrowse:REFRESHALL(),oBrowse:FORCESTABLE,SETAPPFOCUS(oBrowse:getcolumn(1)),.t.}
@ 2,30 DCSAY 'Up ID' get cUpid VALID{|| _LOOK4upid(cUpid),oBrowse:REFRESHALL(),oBrowse:FORCESTABLE,SETAPPFOCUS(oBrowse:getcolumn(1)),.t.}
@ 4,1 DCSAY 'Salesperson' get cSm VALID{|| _LOOK4salesman(cSm),oBrowse:REFRESHALL(),oBrowse:FORCESTABLE,SETAPPFOCUS(oBrowse:getcolumn(1)),.t.}
@ 4,30 DCSAY 'Date in' get dDate VALID{|| _LOOK4date(dDate),oBrowse:REFRESHALL(),oBrowse:FORCESTABLE,SETAPPFOCUS(oBrowse:getcolumn(1)),.t.}

/* ----- Create browse ----- */

@ 8,1 DCBROWSE oBrowse ALIAS 'HNEWUPS'                 ;
	SIZE 110,20                                     ;
	PRESENTATION aPres;
	FREEZELEFT {1,2,3} ;
	THUMBLOCK 500 ;
	sortscolor 0,2;
	SORTUCOLOR 1,6

DCBROWSECOL FIELD HNEWUPS->UPID ;
	HEADER "*UPID/PHN#" PARENT oBrowse  ;
	WIDTH 7    ;
	SORT{|| HNEWUPS->(ORDSETFOCUS('HNEWUPIN')),HNEWUPS->(DBGOTOP()), OBROWSE:REFRESHALL(),oBrowse:forcestable()}

DCBROWSECOL FIELD HNEWUPS->SALESMAN ;
	HEADER "*SALESMAN" PARENT oBrowse  ;
	WIDTH 3      ;
	 SORT{|| HNEWUPS->(ORDSETFOCUS('HNEWUPSM')),HNEWUPS->(DBGOTOP()), OBROWSE:REFRESHALL(),oBrowse:forcestable()}
DCBROWSECOL FIELD HNEWUPS->LASTNAME                     ;
	HEADER "*LASTNAME" HCOLOR 9,3 PARENT oBrowse     ;
	WIDTH 15       ;
	SORT{|| HNEWUPS->(ORDSETFOCUS('HNEWUPNM')),HNEWUPS->(DBGOTOP()), OBROWSE:REFRESHALL(),oBrowse:forcestable()}
DCBROWSECOL FIELD HNEWUPS->FIRSTNAME ;
	HEADER "FIRSTNAME" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD HNEWUPS->DATEIN ;
	HEADER "*DATEIN" PARENT oBrowse  ;
	WIDTH 8     ;
	SORT{|| HNEWUPS->(ORDSETFOCUS('HNEWUPDT')),HNEWUPS->(DBGOTOP()),OBROWSE:REFRESHALL(),oBrowse:forcestable()}
DCBROWSECOL FIELD HNEWUPS->STREET ;
	HEADER "STREET" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD HNEWUPS->CITYSTATE ;
	HEADER "CITYSTATE" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD HNEWUPS->ZIPCODE ;
	HEADER "ZIPCODE" PARENT oBrowse  ;
	WIDTH 5
DCBROWSECOL FIELD HNEWUPS->AREA ;
	HEADER "AREACODE" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD HNEWUPS->UPID ;
	HEADER "PHONE" PARENT oBrowse  ;
	WIDTH 7
DCBROWSECOL FIELD HNEWUPS->WORKPHN ;
	HEADER "WORKPHN" PARENT oBrowse  ;
	WIDTH 12
DCBROWSECOL FIELD HNEWUPS->cellphone ;
	HEADER "CELLPHN" PARENT oBrowse  ;
	WIDTH 10

DCBROWSECOL FIELD HNEWUPS->EMAIL ;
	HEADER "EMAIL" PARENT oBrowse  ;
	WIDTH 35
DCBROWSECOL FIELD HNEWUPS->beback1 ;
	HEADER "Beback1" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD HNEWUPS->beback2 ;
	HEADER "Beback2" PARENT oBrowse  ;
	WIDTH 8

DCBROWSECOL FIELD HNEWUPS->NEWUSED ;
	HEADER "NEWUSED" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD HNEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD HNEWUPS->INTEREST ;
	HEADER "INTEREST" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD HNEWUPS->SALECODE ;
	HEADER "SALECODE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD HNEWUPS->TOMGR ;
	HEADER "TOMGR" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD HNEWUPS->COMMENTS ;
	HEADER "COMMENT" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD HNEWUPS->COMMENT2 ;
	HEADER "COMMENT2" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD HNEWUPS->ADSOURCE ;
	HEADER "ADSOURCE" PARENT oBrowse  ;
	WIDTH 4
DCBROWSECOL FIELD HNEWUPS->DOWNDEAL ;
	HEADER "DOWN" PARENT oBrowse  ;
	WIDTH 1

	DCREAD GUI ;
	FIT ;
	EVAL {||SETAPPFOCUS(obROWSE:GETCOLUMN(1))};
	BUTTONS DCGUI_BUTTON_EXIT
ReTURN nil




CLOSE ALL
ReTURN nil

static function _Look4name()
	local getlist:={}, getoptions , cName:=space(10)
	DCGEToptions TABSTOP saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	@ 10,0 DCSAY 'Enter Name to Search For.' get cName picture '@!'
	DCREAD GUI MODAL ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Name Search' EVAL{|o| SETAPPWINDOW(o)}
	HNEWUPS->(ORDSETFOCUS('HNEWUPNM'))
	SET SOFTSEEK ON
	HNEWUPS->(DBGOTOP())
	SEEK cName
	SET SOFTSEEK OFF
RETURN NIL

static function _Look4upid()
	local getlist:={}, getoptions , cUpId:=space(7)
	DCGEToptions TABSTOP saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	@ 10,0 DCSAY 'Enter Phone to Search For.' get cUpid
	DCREAD GUI MODAL ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Name Search' EVAL{|o| SETAPPWINDOW(o)}

	HNEWUPS->(ORDSETFOCUS('HNEWUPIN'))
	SET SOFTSEEK ON
	HNEWUPS->(DBGOTOP())
	SEEK cUpid
	SET SOFTSEEK OFF
RETURN NIL

static function _Look4date()
	local getlist:={}, getoptions , dDate:=ctod('  /  /  ')
	DCGEToptions TABSTOP saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	@ 10,0 DCSAY 'Enter Date to Search For.' get dDate
	DCREAD GUI MODAL ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Name Search' EVAL{|o| SETAPPWINDOW(o)}

	HNEWUPS->(ORDSETFOCUS('HNEWUPDT'))
	SET SOFTSEEK ON
	HNEWUPS->(DBGOTOP())
	SEEK dDate
	RETURN NIL
static function _Look4salesman()
	local getlist:={}, getoptions , cSM:=space(3)
	DCGEToptions TABSTOP saywidth 0 sayfont '10.Courier New' getfont '10.Courier New' autoresize
	@ 10,0 DCSAY 'Enter SalesPerson to Search For.' get cSM picture '@!'
	DCREAD GUI MODAL ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'Name Search' EVAL{|o| SETAPPWINDOW(o)}

	HNEWUPS->(ORDSETFOCUS('HNEWUPSM'))
	SET SOFTSEEK ON
	HNEWUPS->(DBGOTOP())
	SEEK cSm
	SET SOFTSEEK OFF
	RETURN NIL

FUNCTION HUPLOOK(nRec)
	local getlist:={}
	local getoptions , cprogram , xStatus
	DCGETOPTIONS SAYWIDTH 0 TABSTOP SAYFONT '10.Courier New' GETFONT '10.Courier New' AUTORESIZE

	SELECT HNEWUPS
	HNEWUPS->(DBGOTO(nRec))
	@ 2,5 DCSAY 'HISTORICAL UP INQUIRY'
	@ 3,5 DCSAY 'DATE IN ' + DTOC(HNEWUPS->DATEIN)
	@ 3,20 DCSAY 'SOLDCODE ' + STR(HNEWUPS->SALECODE)
	@ 3,35 DCSAY 'DOWNDEAL ' + STR(HNEWUPS->DOWNDEAL)
	@ 4,5 DCSAY 'PHONE#/UPID ' + HNEWUPS->UPID
	@ 4,30 DCSAY 'AREA CODE ' + HNEWUPS->AREA
	@ 5,5 DCSAY 'FIRSTNAME ' + HNEWUPS->FIRSTNAME  SAYCOLOR 10,0
	@ 6,5 DCSAY 'LASTNAME  ' + HNEWUPS->LASTNAME  SAYCOLOR 10,0
	@ 7,5 DCSAY 'STREET   ' + HNEWUPS->STREET  SAYCOLOR 9,0
	@ 7,50 DCSAY 'ZIPCODE' + HNEWUPS->ZIPCODE  SAYCOLOR 9,0
	@ 10,5 DCSAY 'EMAIL ' + HNEWUPS->EMAIL SAYCOLOR 2,0
	@ 10,65 DCSAY 'WORKPHONE '+  HNEWUPS->WORKPHN  SAYCOLOR 2,0
	@ 11,5 DCSAY 'INTEREST '+  HNEWUPS->INTEREST SAYCOLOR 3,15
	@ 12,5 DCSAY 'UPTYPE '+STR(HNEWUPS->UPTYPE) SAYCOLOR 3,15
	@ 13,5 DCSAY 'N FOR NEW OR U FOR USED '+  HNEWUPS->NEWUSED SAYCOLOR 3,15
	@ 13,50 DCSAY 'AD SOURCE ' + HNEWUPS->ADSOURCE SAYCOLOR 3,15
	@ 14,5 DCSAY 'COMMENTS '+  HNEWUPS->COMMENTS  SAYCOLOR 9,0
	@ 15,5 DCSAY 'MORE COMMENTS '+ HNEWUPS->COMMENT2  SAYCOLOR 9,0
	@ 16,5 DCSAY 'T.O.MANAGER '+ HNEWUPS->TOMGR  SAYCOLOR 10,0
	@ 17,0 DCSAY 'STK# QUOTED ' + HNEWUPS->STKQUOTED +' '+HNEWUPS->YRQUOTED+' '+HNEWUPS->VEHQUOTED
	@ 18,5 DCSAY 'PRICE QUOTED' + STR(HNEWUPS->QUOTED) + ' '+'SALES TAX '+STR(HNEWUPS->TAX)
	@ 19,0 DCSAY 'DOWNPAY '+STR(HNEWUPS->DOWNPAY) + ' '+'REBATE ' +STR(HNEWUPS->REBATE)+' '+'LIEN '+STR(HNEWUPS->LIEN)
	@ 20,0 DCSAY 'AMOUNT FINANCED' + STR(HNEWUPS->FINANCE)
	@ 21,0 DCSAY 'PAYMENTS 24/36/48/60' +' '+STR(HNEWUPS->PAY24)+' '+STR(HNEWUPS->PAY36)+' '+STR(HNEWUPS->PAY48)+' '+STR(HNEWUPS->PAY60)
	@ 22,0 DCSAY 'LEASE QUOTED ' +STR(HNEWUPS->LPAYMENT)+' '+'TERM '+STR(HNEWUPS->LTERM)
	@ 23,0 DCSAY 'NOTES'
	@ 23,6 DCSAY HNEWUPS->NOTES

	DCREAD GUI TO XSTATUS APPWINDOW MAINWINDOW():DRAWINGAREA  FIT ENTEREXIT ADDBUTTONS OPTIONS GETOPTIONS TITLE 'UP RECORD MAINTENANCE'


RETURN NIL

FUNCTION UBROWBYNM()
	LOCAL GETLIST:={} , aotoolbar, oBrowse  , dDatex  , xStatus,oRecord
	LOCAL cLastName:=SPACE(25)  ,cProgram  , aPres ,GETOPTIONS
	cProgram:='BROWBYNM'
	IF.NOT.SECURITY(@cProgram)
	 DC_WINALERT('SECURITY VIOLATION')
    RETURN NIL
   ENDIF

	IF !UseDb({'NEWUPS'})
	 DC_WINALERT('Cannot Open Files     .')
   	 RETURN NIL
   ENDIF

	NEWUPS->(ORDSETFOCUS('NEWUPNM'))
DCGETOPTIONS  ;
	SAYWIDTH 0 ;
	AUTORESIZE
XSTATUS:=.T.
dDatex:=(DATE()-60)


@ 10,0 DCSAY 'ENTER THE LASTNAME YOU WISH TO BROWSE' GET cLastName SAYCOLOR 3,0 SAYFONT '10.COURIER' GETFONT '12.COURIER'
@ 12,0 DCSAY 'ENTER THE DATE TO GO BACK TO BEGIN BROWSE' GET dDatex SAYCOLOR 9,0 SAYFONT '10.COURIER' GETFONT '12.COURIER'
DCREAD GUI TO XSTATUS ENTEREXIT FIT OPTIONS GETOPTIONS TITLE 'ENTER LASTNAME TO BEGIN BROWSE'
IF !XSTATUS
	CLOSE DATABASES
RETURN NIL
ENDIF
SELECT NEWUPS
SET SOFTSEEK ON
SEEK UPPER(cLastName)
SET SOFTSEEK OFF
SET FILTER TO NEWUPS->DATEIN >=dDatex
aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
	 { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

@ 1,1 DCTOOLBAR AoToolBar                               ;
	SIZE 110, 1.5

DCADDBUTTON CAPTION 'Desking Tool;Quote '                              ;
	ACTION  {||	UPQUOTE(NEWUPS->(RECNO())),DC_GETREFRESH(GETLIST)};
	PARENT AoToolBar


DCADDBUTTON CAPTION '2. New Vehicle;Browser  '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("2")};
	COLOR {|| {GRA_CLR_DARKRED,GRA_CLR_WHITE}}  ;
	ACTION {||SBROWBYDESC(),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar                                     ;
	STATIC

DCADDBUTTON CAPTION '3. Used Vehicle;Browser '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("3")};
	COLOR {|| {GRA_CLR_DARKRED,GRA_CLR_WHITE}}  ;
	ACTION {||SBROWBYDESC(),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar                                     ;
	STATIC

DCADDBUTTON CAPTION '5. Payment;Calculator '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("5")};
	COLOR {|| {GRA_CLR_DARKRED,GRA_CLR_WHITE}}  ;
	ACTION {|| QUICK(0,0,0,0,0,0,0,0,0,0,.T.,0),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar                                     ;
	STATIC

DCADDBUTTON CAPTION '6. COMMENTS '                              ;
	SIZE 15                                              ;
	ACCELKEY {ASC("6")};
	COLOR {|| {GRA_CLR_DARKRED,GRA_CLR_WHITE}}  ;
	ACTION {|| UPNOTES(),DC_GETREFRESH(GETLIST)  };
	PARENT AoToolBar                                     ;
	STATIC

/* ----- Create browse ----- */

@ 3,1 DCBROWSE oBrowse ALIAS 'NEWUPS'                 ;
	SIZE 110,20                                     ;
	PRESENTATION aPres;
	FREEZELEFT {1,2,3}

DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "UPID/PHN#" PARENT oBrowse  ;
	WIDTH 7
DCBROWSECOL FIELD NEWUPS->SALESMAN ;
	HEADER "SALESMAN" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->LASTNAME                     ;
	HEADER "LASTNAME" HCOLOR 9,3 PARENT oBrowse     ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->FIRSTNAME ;
	HEADER "FIRSTNAME" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->DATEIN ;
	HEADER "DATEIN" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->STREET ;
	HEADER "STREET" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->CITYSTATE ;
	HEADER "CITYSTATE" PARENT oBrowse  ;
	WIDTH 15
DCBROWSECOL FIELD NEWUPS->ZIPCODE ;
	HEADER "ZIPCODE" PARENT oBrowse  ;
	WIDTH 5
DCBROWSECOL FIELD NEWUPS->AREA ;
	HEADER "AREACODE" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->UPID ;
	HEADER "PHONE" PARENT oBrowse  ;
	WIDTH 7
DCBROWSECOL FIELD NEWUPS->WORKPHN ;
	HEADER "WORKPHN" PARENT oBrowse  ;
	WIDTH 12
DCBROWSECOL FIELD NEWUPS->cellphone ;
	HEADER "CELLPHN" PARENT oBrowse  ;
	WIDTH 10

DCBROWSECOL FIELD NEWUPS->EMAIL ;
	HEADER "EMAIL" PARENT oBrowse  ;
	WIDTH 35
DCBROWSECOL FIELD NEWUPS->beback1 ;
	HEADER "Beback1" PARENT oBrowse  ;
	WIDTH 8
DCBROWSECOL FIELD NEWUPS->beback2 ;
	HEADER "Beback2" PARENT oBrowse  ;
	WIDTH 8

DCBROWSECOL FIELD NEWUPS->NEWUSED ;
	HEADER "NEWUSED" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->UPTYPE ;
	HEADER "TYPE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->INTEREST ;
	HEADER "INTEREST" PARENT oBrowse  ;
	WIDTH 10
DCBROWSECOL FIELD NEWUPS->SALECODE ;
	HEADER "SALECODE" PARENT oBrowse  ;
	WIDTH 1
DCBROWSECOL FIELD NEWUPS->TOMGR ;
	HEADER "TOMGR" PARENT oBrowse  ;
	WIDTH 3
DCBROWSECOL FIELD NEWUPS->COMMENTS ;
	HEADER "COMMENT" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->COMMENT2 ;
	HEADER "COMMENT2" PARENT oBrowse  ;
	WIDTH 20
DCBROWSECOL FIELD NEWUPS->ADSOURCE ;
	HEADER "ADSOURCE" PARENT oBrowse  ;
	WIDTH 4
DCBROWSECOL FIELD NEWUPS->DOWNDEAL ;
	HEADER "DOWN" PARENT oBrowse  ;
	WIDTH 1


DCREAD GUI ;
	FIT ;
	BUTTONS DCGUI_BUTTON_EXIT
CLOSE ALL
ReTURN nil
/*
FUNCTION UPLOOK(lReshow)
	local getlist:={} , getoptions , xstatus ,clPrintUpSheet
	DCGETOPTIONS AUTORESIZE SAYWIDTH 0
	SELECT NEWUPS
	clPrintUpSheet:='N'
	@ 2,5 DCSAY 'HISTORICAL UP INQUIRY'
	@ 3,5 DCSAY 'DATE IN ' + DTOC(NEWUPS->DATEIN)
	@ 3,20 DCSAY 'SOLDCODE ' + STR(NEWUPS->SALECODE)
	@ 3,35 DCSAY 'DOWNDEAL ' + STR(NEWUPS->DOWNDEAL)
	@ 4,5 DCSAY 'PHONE#/UPID ' + NEWUPS->UPID
	@ 4,30 DCSAY 'AREA CODE ' + NEWUPS->AREA
	@ 5,5 DCSAY 'FIRSTNAME ' + NEWUPS->FIRSTNAME  SAYCOLOR 10,0
	@ 6,5 DCSAY 'LASTNAME  ' + NEWUPS->LASTNAME  SAYCOLOR 10,0
	@ 7,5 DCSAY 'STREET   ' + NEWUPS->STREET  SAYCOLOR 9,0
	@ 7,50 DCSAY 'ZIPCODE' + NEWUPS->ZIPCODE  SAYCOLOR 9,0
	@ 10,5 DCSAY 'EMAIL ' + NEWUPS->EMAIL SAYCOLOR 2,0
	@ 10,65 DCSAY 'WORKPHONE '+  NEWUPS->WORKPHN  SAYCOLOR 2,0
	@ 11,5 DCSAY 'INTEREST '+  NEWUPS->INTEREST SAYCOLOR 3,15
	@ 12,5 DCSAY 'UPTYPE '+STR(NEWUPS->UPTYPE) SAYCOLOR 3,15
	@ 13,5 DCSAY 'N FOR NEW OR U FOR USED '+  NEWUPS->NEWUSED SAYCOLOR 3,15
	@ 13,50 DCSAY 'AD SOURCE ' + NEWUPS->ADSOURCE SAYCOLOR 3,15
	@ 14,5 DCSAY 'COMMENTS '+  NEWUPS->COMMENTS  SAYCOLOR 9,0
	@ 15,5 DCSAY 'MORE COMMENTS '+ NEWUPS->COMMENT2  SAYCOLOR 9,0
	@ 16,5 DCSAY 'T.O.MANAGER '+ NEWUPS->TOMGR  SAYCOLOR 10,0
	@ 17,0 DCSAY 'STK# QUOTED ' + NEWUPS->STKQUOTED +' '+NEWUPS->YRQUOTED+' '+NEWUPS->VEHQUOTED
	@ 18,5 DCSAY 'PRICE QUOTED' + STR(NEWUPS->QUOTED) + ' '+'SALES TAX '+STR(NEWUPS->TAX)
	@ 19,0 DCSAY 'DOWNPAY '+STR(NEWUPS->DOWNPAY) + ' '+'REBATE ' +STR(NEWUPS->REBATE)+' '+'LIEN '+STR(NEWUPS->LIEN)
	@ 20,0 DCSAY 'AMOUNT FINANCED' + STR(NEWUPS->FINANCE)
	@ 21,0 DCSAY 'PAYMENTS 24/36/48/60' +' '+STR(NEWUPS->PAY24)+' '+STR(NEWUPS->PAY36)+' '+STR(NEWUPS->PAY48)+' '+STR(NEWUPS->PAY60)
	@ 22,0 DCSAY 'LEASE QUOTED ' +STR(NEWUPS->LPAYMENT)+' '+'TERM '+STR(NEWUPS->LTERM)
	DCREAD GUI FIT OPTIONS GETOPTIONS ADDBUTTONS TITLE 'CURRENT  UP INQUIRY'
	@ 10,20 DCSAY 'DO YOU WISH TO PRINT THIS RECORD' GET clPrintUpSheet  PICTURE '@! L'
	DCREAD GUI FIT ENTEREXIT OPTIONS GETOPTIONS ADDBUTTONS TITLE 'CURRENT  UP INQUIRY'
	IF lReshow
		oBROWSE:SHOW()
	ENDIF
	IF clPrintUpSheet='Y'
		PRTUPSHT()
	ENDIF

RETURN NIL     */

FUNCTION PRTUPSHT()
	LOCAL oPrinter
	SELECT NEWUPS
	DCPRINT ON SIZE 66,132 TO oPrinter FONT '10.Courier' CANCELENABLE
	IF VALTYPE(OPRINTER) # 'O' .OR. !OPRINTER:LACTIVE
				RETURN NIL
	ENDIF

	@ 1,40 DCPRINT SAY 'CUSTOMER FOLLOW UP FORM'
	@ 3,5 DCPRINT SAY 'DATE IN ' + DTOC(NEWUPS->DATEIN)
	@ 3,30 DCPRINT SAY 'SOLDCODE ' + STR(NEWUPS->SALECODE)
	@ 3,50 DCPRINT SAY 'DOWNDEAL ' + STR(NEWUPS->DOWNDEAL)
	@ 4,5 DCPRINT SAY 'PHONE#/UPID ' + NEWUPS->UPID
	@ 4,40 DCPRINT SAY 'AREA CODE ' + NEWUPS->AREA
	@ 5,5 DCPRINT SAY 'FIRSTNAME ' + NEWUPS->FIRSTNAME
	@ 6,5 DCPRINT SAY 'LASTNAME  ' + NEWUPS->LASTNAME
	@ 7,5 DCPRINT SAY 'STREET   ' + NEWUPS->STREET
	@ 7,50 DCPRINT SAY 'ZIPCODE  ' + NEWUPS->ZIPCODE
	@ 10,5 DCPRINT SAY 'EMAIL ' + NEWUPS->EMAIL
	@ 10,65 DCPRINT SAY 'WORKPHONE '+  NEWUPS->WORKPHN
	@ 11,5 DCPRINT SAY 'INTEREST '+  NEWUPS->INTEREST
	@ 12,5 DCPRINT SAY 'UPTYPE '+STR(NEWUPS->UPTYPE)
	@ 13,5 DCPRINT SAY 'N FOR NEW OR U FOR USED '+  NEWUPS->NEWUSED
	@ 13,50 DCPRINT SAY 'AD SOURCE ' + NEWUPS->ADSOURCE
	@ 14,5 DCPRINT SAY 'COMMENTS '+  NEWUPS->COMMENTS
	@ 15,5 DCPRINT SAY 'MORE COMMENTS '+ NEWUPS->COMMENT2
	@ 16,5 DCPRINT SAY 'T.O.MANAGER '+ NEWUPS->TOMGR
	@ 18,5 DCPRINT SAY 'STK# QUOTED ' + NEWUPS->STKQUOTED +' '+NEWUPS->YRQUOTED+' '+NEWUPS->VEHQUOTED
	@ 19,5 DCPRINT SAY 'PRICE QUOTED' + STR(NEWUPS->QUOTED) + ' '+'SALES TAX '+STR(NEWUPS->TAX+NEWUPS->AFTSALETAX)
	@ 20,5 DCPRINT SAY 'DOWNPAY '+STR(NEWUPS->DOWNPAY) + ' '+'REBATE ' +STR(NEWUPS->REBATE)+' '+'LIEN '+STR(NEWUPS->LIEN)
	@ 21,5 DCPRINT SAY 'AMOUNT FINANCED' + STR(NEWUPS->FINANCE)
	@ 22,5 DCPRINT SAY 'PAYMENTS 24/36/48/60' +' '+STR(NEWUPS->PAY24)+' '+STR(NEWUPS->PAY36)+' '+STR(NEWUPS->PAY48)+' '+STR(NEWUPS->PAY60)
	@ 23,5 DCPRINT SAY 'LEASE QUOTED ' +STR(NEWUPS->LPAYMENT)+' '+'TERM '+STR(NEWUPS->LTERM)
	@ 24,5 DCPRINT SAY 'NOTES'
	@ 25,6 DCPRINT SAY NEWUPS->NOTES
	DCPRINT OFF
RETURN NIL


FUNCTION EditUpMail()
	LOCAL GETLIST:={} , pobrowse ,potoolbar , apres, lUlCase:=.t.,GETOPTIONS
	LOCAL IFSCHED:={GRA_CLR_BLACK,GRA_CLR_WHITE}
	LOCAL IFNOTSCHED:={GRA_CLR_WHITE,GRA_CLR_WHITE}
	LOCAL FILENAMEI
	DCGETOPTIONS  ;
		SAYWIDTH 0 ;
		AUTORESIZE sayfont '10.Courier New' getfont '10.Courier New'
	/// CHECK FOR EXISTANCE AND CREATE IF NEESSARY
	IF !FEXISTS('EUPMAIL.DBF')
		COPY FILE (GETFLG('ACCTPATH')+'UPMAIL.DBF') TO (GETFLG('ACCTPATH')+'EUPMAIL.DBF')
		IF USEDB({'EUPMAIL'},1,.F.) // NEW EXCLUSIVE
		 ZAP
		 INDEX ON EUPMAIL->LASTNAME TO EUPMAIL
		 EUPMAIL->(DBCLOSEAREA())
		ELSE
        AboCloseAll()
		  RETURN NIL
		ENDIF
	ENDIF

	if use_udf('Upmail',.t.)
		goto top
	else
		close all
		return nil
	endif

	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
	 { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */

	@ 1,0 DCCHECKBOX lUlCase PROMPT 'Convert to Upper/Lower Case When Done'

	@ 3,1 DCTOOLBAR PoToolBar                               ;
		SIZE 50, 2  BUTTONSIZE 25,2

	DCADDBUTTON CAPTION 'DELETE RECORD'                              ;
		ACTION {|| _DELupMAIL(), DC_GetRefresh(GetList),Pobrowse:forcestable()}        ;
		PARENT PoToolBar                                     ;
		TOOLTIP 'DELETE HIGHLIGHTED RECORD'

	DCADDBUTTON CAPTION 'Create eMail File'                              ;
		ACTION {|| _MakeUpeMAIL(), DC_GetRefresh(GetList),Pobrowse:forcestable()}        ;
		PARENT PoToolBar                                     ;
		TOOLTIP 'DELETE HIGHLIGHTED RECORD'
	/*
	DCADDBUTTON CAPTION 'Send Mass eMail'                              ;
		ACTION {|| UPMASSIVEEMAIL(), DC_GetRefresh(GetList),Pobrowse:forcestable()}        ;
		PARENT PoToolBar                                     ;
		TOOLTIP 'Select and Send eMail' */



/* ----- Create browse ----- */

	@ 5,1 DCBROWSE PoBrowse ALIAS 'UPMAIL'                 ;
		SIZE 110,20                                     ;
		PRESENTATION aPres;
		FREEZELEFT{1,2} ;
		COLOR {||IF(!DELETED(),IFSCHED,IFNOTSCHED)} ;
		EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITEXIT


	DCBROWSECOL FIELD upmail->LASTNAME                     ;
		HEADER "LASTNAME" HCOLOR 1,5 PARENT PoBrowse      ;
		WIDTH 20
	DCBROWSECOL FIELD upmail->FIRSTNAME                     ;
		HEADER "FIRSTNAME" HCOLOR 1,5 PARENT PoBrowse    ;
		WIDTH 20
	DCBROWSECOL FIELD upmail->STREET                     ;
		HEADER "STREET" HCOLOR 1,5 PARENT PoBrowse    ;
		WIDTH 20
	DCBROWSECOL FIELD upmail->CITYSTATE                     ;
		HEADER "CITYSTATE" HCOLOR 1,5 PARENT PoBrowse    ;
		WIDTH 20
	DCBROWSECOL FIELD upmail->ZIPcode                     ;
		HEADER "ZIP" HCOLOR 1,5 PARENT PoBrowse    ;
		WIDTH 5
	DCBROWSECOL FIELD upmail->Email                     ;
		HEADER "eMail" HCOLOR 1,5 PARENT PoBrowse      ;
		WIDTH 30

	DCBROWSECOL FIELD upmail->Salesman                     ;
		HEADER "SALESPERSON" HCOLOR 1,3 PARENT PoBrowse     ;
		WIDTH 2

	DCBROWSECOL FIELD upmail->datein                     ;
		HEADER "Date in" HCOLOR 1,3 PARENT PoBrowse     ;
		WIDTH 8
	DCBROWSECOL FIELD upmail->interest                     ;
		HEADER "Interest" HCOLOR 1,5 PARENT PoBrowse    ;
		WIDTH 10

	DCREAD GUI ;
		EVAL{||SETAPPFOCUS(poBrowse:getcolumn(1))};
		FIT ;
		TITLE 'Up mail Browser'  ;
		OPTIONS GETOPTIONS ;
		BUTTONS DCGUI_BUTTON_EXIT


	/// convert to U/L case
	IF lUlcase
   	uconuptolow()
	ENDIF
	CLOSE DATABASES


RETURN NIL

STATIC FUNCTION _MakeUpeMail()
	LOCAL nEmail:=0

	IF !UseDb({'EUPMAIL'},1,.F.)
		 DC_WINALERT('Cannot Open Files     .')
	   	 RETURN NIL
	ENDIF

	SELECT UPMAIL
	UPMAIL->(DBGOTOP())
	DO WHILE !UPMAIL->(EOF())
		  IF !EMPTY(UPMAIL->EMAIL)
			SELECT EUPMAIL
			IF EUPMAIL->(DC_ADDREC())
				REPLACE EUPMAIL->LASTNAME WITH UPMAIL->LASTNAME
				REPLACE EUPMAIL->FIRSTNAME WITH UPMAIL->FIRSTNAME
				REPLACE EUPMAIL->STREET WITH UPMAIL->STREET
				REPLACE EUPMAIL->CITYSTATE WITH UPMAIL->CITYSTATE
				REPLACE EUPMAIL->ZIPCODE WITH UPMAIL->ZIPCODE
				REPLACE EUPMAIL->SALESMAN WITH UPMAIL->SALESMAN
				REPLACE EUPMAIL->DATEIN WITH UPMAIL->DATEIN
				REPLACE EUPMAIL->EMAIL WITH UPMAIL->EMAIL
				REPLACE EUPMAIL->INTEREST WITH UPMAIL->INTEREST
				REPLACE EUPMAIL->UPTYPE WITH UPMAIL->UPTYPE
			ENDIF
		  ENDIF
		  SKIP ALIAS UPMAIL
	ENDDO
	EUPMAIL->(DBCOMMIT())
	// NOW CLEAR BAD EMAILS

	SELECT EUPMAIL
	EUPMAIL->(DBGOTOP())
	DO WHILE !EUPMAIL->(EOF())
		IF EUPMAIL->email='WNG@A.COM'
			EUPMAIL->(DBDELETE())
		ENDIF
		IF EUPMAIL->email='WNG'
			EUPMAIL->(DBDELETE())
		ENDIF
		IF EUPMAIL->email='AFL'
			EUPMAIL->(DBDELETE())
		ENDIF
		IF EUPMAIL->email='DNH'
			EUPMAIL->(DBDELETE())
		ENDIF

		IF EUPMAIL->email='ASKFORLATER@A.COM'
			EUPMAIL->(DBDELETE())
		ENDIF
		IF EUPMAIL->email='DNH@A.COM'
			EUPMAIL->(DBDELETE())
		ENDIF
		/// replace email adresses that do not have a @ in them or have a.com
		if !empty(EUPMAIL->email)
			nEmail:=0
			nEmail:=AT('@',EUPMAIL->email,2)
			if nEmail=0
			  EUPMAIL->(DBDELETE())
			ENDIF
			IF nEmail > 0
				if substr(EUPMAIL->email,nEmail,6)='@A.COM'
				  EUPMAIL->(DBDELETE())
				ENDIF
				if substr(EUPMAIL->email,nEmail,6)='@a.com'
				 EUPMAIL->(DBDELETE())
				ENDIF
				if substr(EUPMAIL->email,nEmail,5)='@.COM'
				  EUPMAIL->(DBDELETE())
				ENDIF
				if substr(EUPMAIL->email,nEmail,5)='@.com'
				  EUPMAIL->(DBDELETE())
				ENDIF
		   ENDIF
		 ENDIF
		 SKIP ALIAS EUPMAIL
	ENDDO
	SELECT EUPMAIL
	PACK
	EUPMAIL->(DBCLOSEAREA())
	IF !UseDb({'EUPMAIL'},1,.F.,.T.)   /// REOPEN and REINDEX
		 DC_WINALERT('Cannot Open Files     .')
	   	 RETURN NIL
	ENDIF



	DC_WINALERT('File eUpMail.dbf has been created')
	/// NOW SHOW EUPMAIL BEFORE RUNNING
	EditUpEMail()

	RETURN NIL


function EditUpEMail(lOpenfile)
	LOCAL GETLIST:={} , obrowse ,otoolbar , apres, lUlCase:=.t.  ,GETOPTIONS
	LOCAL IFSCHED:={GRA_CLR_BLACK,GRA_CLR_WHITE}
	LOCAL IFNOTSCHED:={GRA_CLR_WHITE,GRA_CLR_WHITE}
	DEFAULT lOpenfile:=.f.
	DCGETOPTIONS  ;
		SAYWIDTH 0 ;
		AUTORESIZE sayfont '10.Courier New' getfont '10.Courier New'

	IF lOpenfile

		IF !UseDb({'EUPMAIL'},1,.F.)
		 DC_WINALERT('Cannot Open Files     .')
	    RETURN NIL
		ENDIF
	ENDIF


	aPres := ;
  { { XBP_PP_COL_HA_FGCLR, GRA_CLR_WHITE },    /*  Header FG Color  */     ;
    { XBP_PP_COL_HA_BGCLR, GRA_CLR_DARKGRAY }, /*  Header BG Color  */     ;
    { XBP_PP_COL_DA_ROWSEPARATOR, XBPCOL_SEP_DOTTED },  /* Row Sep  */     ;
	 { XBP_PP_COL_DA_HILITE_BGCLR, GRA_CLR_PALEGRAY }, /*  HILITE BG Color  */     ;
	 { XBP_PP_COL_DA_COLSEPARATOR, XBPCOL_SEP_DOTTED },  /* Col Sep  */     ;
    { XBP_PP_COL_DA_ROWHEIGHT, 20 },                  /* Row Height */     ;
    { XBP_PP_COL_DA_CELLHEIGHT, 20 }  }              /* Cell Height */

/* ----- Create ToolBar ----- */


	@ 3,1 DCTOOLBAR oToolBar                               ;
		SIZE 50, 2  BUTTONSIZE 25,2

	DCADDBUTTON CAPTION 'DELETE RECORD'                              ;
		ACTION {|| _DELEUPMAIL(), DC_GetRefresh(GetList),obrowse:forcestable()}        ;
		PARENT oToolBar                                     ;
		TOOLTIP 'DELETE HIGHLIGHTED RECORD'

	DCADDBUTTON CAPTION 'Send Mass eMail'                              ;
		ACTION {|| UPMASSIVEEMAIL(), DC_GetRefresh(GetList),obrowse:forcestable()}        ;
		PARENT oToolBar                                     ;
		TOOLTIP 'Select and Send eMail'



/* ----- Create browse ----- */

	@ 5,1 DCBROWSE oBrowse ALIAS 'EUPMAIL'                 ;
		SIZE 110,20                                     ;
		PRESENTATION aPres;
		FREEZELEFT{1,2} ;
		COLOR {||IF(!DELETED(),IFSCHED,IFNOTSCHED)} ;
		EDIT xbeBRW_ItemSelected MODE DCGUI_BROWSE_EDITEXIT


	DCBROWSECOL FIELD EUPMAIL->Email                     ;
		HEADER "eMail" HCOLOR 1,5 PARENT oBrowse      ;
		WIDTH 20

	DCBROWSECOL FIELD EUPMAIL->LASTNAME                     ;
		HEADER "LASTNAME" HCOLOR 1,5 PARENT oBrowse      ;
		WIDTH 20
	DCBROWSECOL FIELD EUPMAIL->FIRSTNAME                     ;
		HEADER "FIRSTNAME" HCOLOR 1,5 PARENT oBrowse    ;
		WIDTH 20
	DCBROWSECOL FIELD EUPMAIL->STREET                     ;
		HEADER "STREET" HCOLOR 1,5 PARENT oBrowse    ;
		WIDTH 20
	DCBROWSECOL FIELD EUPMAIL->CITYSTATE                     ;
		HEADER "CITYSTATE" HCOLOR 1,5 PARENT oBrowse    ;
		WIDTH 20
	DCBROWSECOL FIELD EUPMAIL->ZIPcode                     ;
		HEADER "ZIP" HCOLOR 1,5 PARENT oBrowse    ;
		WIDTH 5

	DCBROWSECOL FIELD EUPMAIL->Salesman                     ;
		HEADER "SALESPERSON" HCOLOR 1,3 PARENT oBrowse     ;
		WIDTH 2

	DCBROWSECOL FIELD EUPMAIL->datein                     ;
		HEADER "Date in" HCOLOR 1,3 PARENT oBrowse     ;
		WIDTH 8
	DCBROWSECOL FIELD EUPMAIL->interest                     ;
		HEADER "Interest" HCOLOR 1,5 PARENT oBrowse    ;
		WIDTH 10

	DCREAD GUI ;
		EVAL{||SETAPPFOCUS(oBrowse:getcolumn(1))};
		FIT ;
		TITLE 'uP eMail Browser' ;
		OPTIONS GETOPTIONS ;
		BUTTONS DCGUI_BUTTON_EXIT

	SELECT EUPMAIL
	PACK
	EUPMAIL->(DBCLOSEAREA())

	RETURN NIL







STATIC FUNCTION _DelUpMail()
	SELECT UPMAIL
	UPMAIL->(BV_BLANK(.T.,.F.))
   RETURN NIL

STATIC FUNCTION _DelEUPMAIL()
	EUPMAIL->(BV_BLANK(.T.,.F.))
RETURN NIL



Function Uconuptolow()
	local cFirstname, cLastName , cStreet,cCity ,oDialog
	SELECT UPMAIL
	ORDSETFOCUS()
	GOTO TOP
	ODIALOG:=DC_WAITON('Now Converting to U/L case')
	DO WHILE !UPMAIL->(EOF())
		cFirstname:=UPMAIL->FIRSTNAME
		cLastName:=UPMAIL->lastname
		cStreet:=UPMAIL->street
		cCity:=UPMAIL->citystate
		UxuPTOLOW(@cFirstname)
		UxUPTOLOW(@cLastName)
		UxUPTOLOW(@cStreet)
		UxUPTOLOW(@cCity)
		REPLACE UPMAIL->FIRSTNAME WITH cFirstname
		REPLACE UPMAIL->lastNAME WITH cLastName
		REPLACE UPMAIL->street WITH cStreet
		REPLACE UPMAIL->citystate WITH cCity
		SKIP ALIAS UPMAIL
	ENDDO
	UPMAIL->(DBCOMMIT())
	dc_impl(odialog)
	close all
	dc_winalert('File Converted')
return nil


FUNCTION UUPTOLOW(cString)
	local Cstringx , n,i
	Cstringx:=substr(cString,1,1)
	n:=2
	for I=1 to (len(cstring)-1)
		if !empty(substr(cstring,n,1))
			cStringx:=Cstringx+lower(substr(cString,n,1))
		endif
		if substr(cstring,n)=' '
			cstring:=cstringx
			return nil
		endif
		n:=n+1
	next
	cstring:=cstringx
	//DC_WINALERT(CSTRINGI)
return cstring

FUNCTION UxUPTOLOW(cString)
	local Cstringx , i , n
	LOCAL newx:=.f.
	Cstringx:=substr(cString,1,1)
	CSTRINGx:=UPPER(CSTRINGx)
	n:=2
	for I=1 to (len(cstring)-1)
		if !empty(substr(cstring,n,1))
			if !newx
				cStringx:=Cstringx+lower(substr(cString,n,1))
			else
				cStringx:=Cstringx+upper(substr(cString,n,1))
				newx:=.f.
			endif
		endif
		if substr(cstring,n,1)=' '                /// start caps for new word
			CSTRINGx:=CSTRINGx+' '
			newx:=.t.
		endif
		if substr(cstring,n,1)="-"                /// start caps for new word
			newx:=.t.
		endif
		if substr(cstring,n,1)="."                /// start caps for new word
			newx:=.t.
		endif
		if substr(cstring,n,1)="/"                /// start caps for new word
			newx:=.t.
		endif
		if substr(cstring,n,1)="&"
			newx:=.t.
		endif

		n:=n+1
	next
	IF SUBSTR(CSTRINGx,1,2)='Mc'                 /// CONVERT THE MICS
		CSTRINGx:=SUBSTR(CSTRINGx,1,2)+UPPER(SUBSTR(CSTRINGx,3,1))+SUBSTR(CSTRINGx,4,16)
	ENDIF
	IF SUBSTR(CSTRINGx,1,2)="O'"                 /// CONVERT THE OS
		CSTRINGx:=SUBSTR(CSTRINGx,1,2)+UPPER(SUBSTR(CSTRINGx,3,1))+SUBSTR(CSTRINGx,4,16)
	ENDIF
	IF SUBSTR(CSTRINGx,1,2)="D'"                 /// CONVERT THE GINNIES
		CSTRINGx:=SUBSTR(CSTRINGx,1,2)+UPPER(SUBSTR(CSTRINGx,3,1))+SUBSTR(CSTRINGx,4,16)
	ENDIF

	cstring:=cstringx


	//DC_WINALERT(CSTRINGI)
return cstring



















