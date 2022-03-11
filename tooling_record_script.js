const g_Color_Default		= "#BFBFBF";
const g_Color_White			= "#FFFFFF";
const g_Color_Btn			= "#A6A6A6";
const g_Color_Black			= "#000000";
const g_Color_Letter1		= "#00B0F5";

// setting for over view
const g_iHeight_Title		= 36;
const g_iHeight_Row			= 18;
const g_iSize_Title			= 18;
const g_iLength_Column 		= 16;
const g_iLength_Row			= 30;
const g_iNumber_Header 		= 2;
const g_iNumber_Row			= g_iNumber_Header + g_iLength_Row;
const g_iNumber_FormRow		= 28;
const g_sDataStart			= "C";

// setting for sheet reset
const g_bResetCellValue 	= false;
const g_bResetFillColor 	= true;
const g_bResetConFormat 	= true;
const g_bRunOverView		= true;
const g_bRunToolingForm		= true;
const g_iRunToolingFormNum	= 0;
const g_sMainSheet			= "OVERVIEW";

let left					= 0;
let righ					= 1;
let g_iBuff_num				= [0,0];
let g_sBuff_chr				= ["",""];
let g_sBuff_rt1				= ["",""];
let g_sBuff_rt2				= ["",""];

let letterBuff1				= "";
let letterBuff2				= "";
let letterBuff3				= "";
let letterBuff4				= "";

Excel.run(async (context) =>
{
	if( g_bRunOverView ) {
		let sheet = context.workbook.worksheets.getItem(g_sMainSheet);

		await context.sync().then(() => {}).catch((err) => {
			if(err) {
				sheet = context.workbook.worksheets.add(g_sMainSheet);
				console.log("**** Sheet " + g_sMainSheet + " has been created ****");
			}
		});
		
		await context.sync();
		sheet.activate();
		sheet.load("name");
		sheet.protection.unprotect("1");
		
		await context.sync();
		BuildToolingOverView(sheet);
	}
	
	if(g_bRunToolingForm)
	{
		if( g_iRunToolingFormNum != 0 )
		{
			sheet = context.workbook.worksheets.getItem( g_iRunToolingFormNum.toString() );
			
			await context.sync().then(() => {}).catch((err) => {
				if(err) {
				sheet = context.workbook.worksheets.add( g_iRunToolingFormNum.toString() );
					console.log(`"**** Sheet ${g_iRunToolingFormNum} has been created ****"`);
				}
			});
			
			await context.sync();
			sheet.activate();
			sheet.load("name");
			sheet.protection.unprotect("1");

			await context.sync();
			BuildToolingForm(sheet);
		}
		else
		{
			for(let i=1; i<=g_iLength_Row; i++)
			{
				sheet = context.workbook.worksheets.getItem(`${i}`);
				
				await context.sync().then(() => {}).catch((err) => {
					if(err) {
						sheet = context.workbook.worksheets.add(`${i}`);
						console.log(`"**** Sheet ${i} has been created ****"`);
					}
				});
				
				await context.sync();
				sheet.activate();
				sheet.load("name");
				sheet.protection.unprotect("1");

				await context.sync();
				BuildToolingForm(sheet);
			}
		}
	}
	await context.sync().then(() => { console.log("======= Building Tooling & Equipment Record Completed ======="); });
})

function ResetSheetDefault(range)
{
	if (g_bResetCellValue) {
		range.clear();
	}
	if (g_bResetFillColor) {
		range.format.fill.clear();
	}
	if (g_bResetConFormat) {
		range.conditionalFormats.clearAll();
	}
}

function BuildToolingOverView(sheet)
{
	console.log(`======= "${sheet.name}" Repair Initialized =======`);
	
	let range = sheet.getRange();
	
	ResetSheetDefault(range);
	
	range.format.rowHeight					= 22;
	range.format.font.size					= 11;
	range.format.protection.locked			= true;
	range.format.protection.formulaHidden	= false;
	range.format.font.name					= "Calibri";
	range.format.horizontalAlignment		= "Center";
	range.format.verticalAlignment			= "Center";
	range.format.font.bold					= false;
	range.format.font.italic				= false;
	
	range = sheet.getRange("B" + (g_iNumber_Header + 1) + ":B" + g_iNumber_Row);
	range.format.horizontalAlignment		= "Left";
	range.format.protection.locked			= false;
	
	
	let sDataEnd = String.fromCharCode(g_sDataStart.charCodeAt(0) + (g_iLength_Column - 1));
	range = sheet.getRange("A1:" + sDataEnd + "1");
	range.format.rowHeight					= g_iHeight_Title;
	range.format.font.size					= g_iSize_Title;
	range.format.protection.locked			= true;
	range.values = "TOOLING & EQUIPMENT IN/OUT RECORD";
	range.merge();
	range.format.font.bold					= true;
	range.format.horizontalAlignment		= "Center";
	range.format.verticalAlignment			= "Center";
		
	range = sheet.getRange("A" + g_iNumber_Header);
	range.values							= "ITEM";
	range.format.columnWidth				= 35;
	
	range = sheet.getRange("B" + g_iNumber_Header);
	range.values							= "DESCRIPTION";
	range.format.columnWidth				= 233.;
	
	range = sheet.getRange(g_sDataStart + g_iNumber_Header + ":" + sDataEnd + g_iNumber_Header);
	range.merge();
	range.values							= "STATUS";
	range.format.columnWidth				= 60.;
	
	range = sheet.getRange("A" + g_iNumber_Header + ":" + sDataEnd + g_iNumber_Header);
	range.format.fill.color					= g_Color_Black;
	range.format.font.color					= g_Color_White;
	range.format.font.bold					= true;
	range.format.horizontalAlignment		= "Center";
	range.format.verticalAlignment			= "Center";
	
	
	let i, j=1;
	for(i=(g_iNumber_Header + 1); i<=(g_iNumber_Header + g_iLength_Row); i++)
	{
		range = sheet.getRange("A" + i);
		range.format.rowHeight = 22;
		range.values = j;
		let hyperlink =
		{
			textToDisplay:j.toString(),
			ScreenTip:"",
			documentReference:"'" + j + "'!A1"
		};
		range.hyperlink = hyperlink;

		let charC="C", char1="", char2="", char3="B";
		for(k=1; k<=g_iLength_Column; k++)
		{
			range = sheet.getRange(charC + i);
			range.formulas = `=IF('` + j + `'!$` + char1 + char2 + char3 + `$2<>"",'` + j + `'!$` + char1 + char2 + char3 + `$3,"")`;
			
			condition = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
			condition.custom.rule.formula = `=AND($` + charC + `$` + i + `>2,$` + charC + `$` + i + `<>"")`;
			condition.custom.format.fill.color = "Green";
			condition = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
			condition.custom.rule.formula = `=$` + charC + `$` + i + `=2`;
			condition.custom.format.fill.color = "Yellow";
			condition = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
			condition.custom.rule.formula = `=$` + charC + `$` + i + `<=1`;
			condition.custom.format.fill.color = "Red";
			charC = String.fromCharCode(charC.charCodeAt(0) + 1);
			char3 = String.fromCharCode(char3.charCodeAt(0) + 2); // this inc need to minus "B"
			if( char3.charCodeAt(0) > 90 ) {
				char3 = "B";
				if( char2 === "" ) {
					char2 = "A";
				}
				else {
					char2 = String.fromCharCode(char2.charCodeAt(0) + 1);
					if( char2.charCodeAt(0) > 90 ) {
						char2 = "A";
						if( char1 === "" ) {
							char1 = "A";
						}
						else {
							char1 = String.fromCharCode(char1.charCodeAt(0) + 1);
						}
					}
				}
			}
		}
		j++;
	}
	sheet.protection.protect({}, "1");
}

function BuildToolingForm(sheet)
{
	console.log(`======= "${sheet.name}" Repair Initialized =======`);
	
	let range = sheet.getRange();
	range.format.rowHeight					= 18;
	range.format.font.size					= 11;
	range.format.protection.locked			= true;
	range.format.protection.formulaHidden	= false;
	range.format.font.name					= "Calibri";
	range.format.horizontalAlignment		= "Center";
	range.format.verticalAlignment			= "Center";
	range.format.font.bold					= false;
	range.format.font.italic				= false;
	sheet.freezePanes.unfreeze();

	ResetLettersBuffer("", "", "", "B");
	letterBuff1 = String.fromCharCode(letterBuff1.charCodeAt(0) + ((g_iLength_Column * 2) - 2));
	IncreaseLetters(letterBuff4, letterBuff3, letterBuff2, letterBuff1);
	
	range = sheet.getRange("A1:" + letterBuff4 + letterBuff3+ letterBuff2 + letterBuff1 + g_iNumber_FormRow);
	range.unmerge();
	
	range = sheet.getRange("A1");
	let hyperlink = {
		textToDisplay:"<< BACK",
		ScreenTip:"",
		documentReference:`'${g_sMainSheet}'!A` + (parseInt(sheet.name) + 2)
	};
	range.hyperlink							= hyperlink;
	range.format.horizontalAlignment		= "Center";
	range.format.verticalAlignment			= "Center";
	range.format.font.size					= 12;
	range.format.font.bold					= true;
	range.format.fill.color					= g_Color_Btn;
	
	range = sheet.getRange("B1:" + letterBuff4 + letterBuff3+ letterBuff2 + letterBuff1 + "1");
	range.format.rowHeight					= g_iHeight_Title;
	range.format.font.size					= g_iSize_Title;
	range.format.protection.locked			= true;
	range.format.font.bold					= true;
	range.format.horizontalAlignment		= "Center";
	range.format.verticalAlignment			= "Center";
	range.merge();
	range.format.font.size					= 18;
	range.format.font.bold					= true;
	range.format.fill.color					= g_Color_Default;
	range.values = `=IF('${g_sMainSheet}'!B` + (parseInt(sheet.name) + 2) + `<>"",'${g_sMainSheet}'!B` + (parseInt(sheet.name) + 2) + `,"EMPTY")`;
	
	range = sheet.getRange("A2");
	range.format.columnWidth				= 140; //12
	range.values							= "DIAMETER";
	range.format.fill.color					= g_Color_Default;
	
	range = sheet.getRange("A3");
	range.values							= "BALANCE";
	range.format.fill.color					= g_Color_Default;
	
	range = sheet.getRange("A4");
	range.values							= "COMMENT/DATE";
	range.format.fill.color					= g_Color_Default;
	
	ResetLettersBuffer("", "", "", "A");
	for(let i=0; i<(g_iLength_Column * 2); i++)
	{
		IncreaseLetters(letterBuff4, letterBuff3, letterBuff2, letterBuff1);
		range = sheet.getRange( letterBuff4 + letterBuff3 + letterBuff2 + letterBuff1 + "4" );
		range.format.fill.color		= g_Color_Default;
		range.format.columnWidth	= 35;
		//range.format.borders.style = "none";
		if(i % 2 === 1 )
		{
			range.values = "OUT";
		}
		else
		{
			range.values = "IN";
		}
	}
	
	g_iBuff_num[left] = 0;
	g_iBuff_num[righ] = 0;
	g_sBuff_chr[left] = "";
	g_sBuff_chr[righ] = "";
	g_sBuff_rt1[left] = "";
	g_sBuff_rt1[righ] = "";
	g_sBuff_rt2[left] = "";
	g_sBuff_rt2[righ] = "";

	
	letterBuff1	= "A"
	letterBuff2	= ""
	for(let i=0; i<g_iLength_Column; i++) {
		GetNextLetters(sheet, letterBuff1);
		letterBuff1 = g_sBuff_chr[righ];
		
		range = sheet.getRange( g_sBuff_rt2[left] + g_sBuff_rt1[left] + g_sBuff_chr[left] + "2:" + g_sBuff_rt2[righ] + g_sBuff_rt1[righ] + g_sBuff_chr[righ] + "2");
		range.merge();
		range.format.protection.locked		= false;
		range.format.fill.color				= g_Color_Default;
		range.format.horizontalAlignment	= "Center";
		range.format.verticalAlignment		= "Center";
		
		range = sheet.getRange( g_sBuff_rt2[left] + g_sBuff_rt1[left] + g_sBuff_chr[left] + "3:" + g_sBuff_rt2[righ] + g_sBuff_rt1[righ] + g_sBuff_chr[righ] + "3");
		range.merge();
		range.format.font.bold				= true;
		range.format.fill.color				= g_Color_Default;
		range.format.horizontalAlignment	= "Center";
		range.format.verticalAlignment		= "Center";
		range.values						= `=IF(${g_sBuff_rt2[left] + g_sBuff_rt1[left] + g_sBuff_chr[left]}2<>"",SUM(${g_sBuff_rt2[left] + g_sBuff_rt1[left] + g_sBuff_chr[left]}5:${g_sBuff_rt2[left] + g_sBuff_rt1[left] + g_sBuff_chr[left] + g_iNumber_FormRow})-SUM(${g_sBuff_rt2[righ] + g_sBuff_rt1[righ] + g_sBuff_chr[righ]}5:${g_sBuff_rt2[righ] + g_sBuff_rt1[righ] + g_sBuff_chr[righ] + g_iNumber_FormRow}),"")`;
		range.select();
		
		condition = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
		condition.custom.rule.formula		= `=AND($` + g_sBuff_rt2[left] + g_sBuff_rt1[left] + g_sBuff_chr[left] + `$3>2,$` + g_sBuff_rt2[left] + g_sBuff_rt1[left] + g_sBuff_chr[left] + `$3<>"")`;
		condition.custom.format.font.color	= "Green";
		
		condition = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
		condition.custom.rule.formula 		= `=$` + g_sBuff_rt2[left] + g_sBuff_rt1[left] + g_sBuff_chr[left] + `$3=2`;
		condition.custom.format.font.color	= "Yellow";
		
		condition = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
		condition.custom.rule.formula 		= `=$` + g_sBuff_rt2[left] + g_sBuff_rt1[left] + g_sBuff_chr[left] + `$3<=1`;
		condition.custom.format.font.color	= "Red";
		
		if(g_sBuff_rt1[righ] !== "")
		{
			g_sBuff_rt1[left] = g_sBuff_rt1[righ];
		}
		
		if(g_sBuff_rt2[righ] !== "")
		{
			g_sBuff_rt2[left] = g_sBuff_rt2[righ];
		}
	}
	
	range = sheet.getRange("A5:" + g_sBuff_rt2[righ] + g_sBuff_rt1[righ] + g_sBuff_chr[righ] + g_iNumber_FormRow);
	range.format.protection.locked = false;
	sheet.freezePanes.freezeColumns(1);
	sheet.freezePanes.freezeRows(4);
	sheet.protection.protect({}, "1");
}

function GetNextLetters(sheet, start_char)
{
	g_iBuff_num[left] = start_char.charCodeAt(0) + 1;
	if(g_iBuff_num[left] > 90)
	{
		g_sBuff_chr[left] = "A";
		g_iBuff_num[left] = g_sBuff_chr[left].charCodeAt(0) + (g_iBuff_num[left] - 90 - 1);
		
		if(g_sBuff_rt1[left] === "")
		{
			g_sBuff_rt1[left] = "A";
		}
		else
		{
			g_sBuff_rt1[left] = String.fromCharCode( g_sBuff_rt1[left].charCodeAt(0) + 1);
			if( g_sBuff_rt1[left].charCodeAt(0) > 90 )
			{
				g_sBuff_rt1[left] = "A";
				if(g_sBuff_rt2[left] === "")
				{
					g_sBuff_rt2[left] = "A";
				}
				else
				{
					g_sBuff_rt2[left] = String.fromCharCode(g_sBuff_rt2[left].charCodeAt(0) + 1);
				}
			}
		}
	}
	
	g_sBuff_chr[left] = String.fromCharCode(g_iBuff_num[left]);
	
	g_iBuff_num[righ] = g_iBuff_num[left] + 1;
	if(g_iBuff_num[righ] > 90)
	{
		g_sBuff_chr[righ] = "A";
		g_iBuff_num[righ] = g_sBuff_chr[righ].charCodeAt(0) + (g_iBuff_num[righ] - 90 - 1);
		g_sBuff_rt1[righ] = g_sBuff_rt1[left];
		
		if(g_sBuff_rt1[righ] === "")
		{
			g_sBuff_rt1[righ] = "A";
		}
		else
		{
			g_sBuff_rt1[righ] = String.fromCharCode(g_sBuff_rt1[righ].charCodeAt(0) + 1);
			if( g_sBuff_rt1[righ].charCodeAt(0) > 90 )
			{
				g_sBuff_rt1[righ] = "A";
				if(g_sBuff_rt2[righ] === "")
				{
					g_sBuff_rt2[righ] = "A";
				}
				else
				{
					g_sBuff_rt2[righ] = String.fromCharCode(g_sBuff_rt2[righ].charCodeAt(0) + 1);
				}
			}
		}
	}
	g_sBuff_chr[righ] = String.fromCharCode(g_iBuff_num[righ]);
}

function IncreaseLetters(charInput4, charInput3, charInput2, charInput1)
{
	let charIn1 = charInput1;
	let charIn2 = charInput2;
	let charIn3 = charInput3;
	let charIn4 = charInput4;
	
	let charIn1 = String.fromCharCode(charIn1.charCodeAt(0) + 1 );
	if( charIn1.charCodeAt(0) > 90 )
	{
		charIn1 = String.fromCharCode( 64 + (charIn1.charCodeAt(0)) - 90 );
		if(charIn2 === "" )
		{
			charIn2 = "A";
		}
		else
		{
			charIn2 = String.fromCharCode( charIn2.charCodeAt(0) + 1 );
			if( charIn2.charCodeAt(0) > 90 )
			{
				charIn2 = String.fromCharCode( 64 + (charIn2.charCodeAt(0)) - 90 );
				if(charIn3 === "" )
				{
					charIn3 = "A";
				}
				else
				{
					charIn3 = String.fromCharCode( charIn2.charCodeAt(0) + 1 );
					if( charIn3.charCodeAt(0) > 90 )
					{
						charIn3 = String.fromCharCode( 64 + (charIn3.charCodeAt(0)) - 90 );
						if(charIn4 === "" )
						{
							charIn4 = "A";
						}
						else
						{
							charIn4 = String.fromCharCode( charIn4.charCodeAt(0) + 1 );
							if( charIn4.charCodeAt(0) > 90 )
							{
								console.log("function 'IncreaseLetters()' has run out of buffers...");
							}
						}
					}
				}
			}
		}
	}
	letterBuff1 = charIn1;
	letterBuff2 = charIn2;
	letterBuff3 = charIn3;
	letterBuff4 = charIn4;
}

function ResetLettersBuffer(charInput4, charInput3, charInput2, charInput1)
{
	letterBuff1 = charInput1;
	letterBuff2 = charInput2;
	letterBuff3 = charInput3;
	letterBuff4 = charInput4;
}

