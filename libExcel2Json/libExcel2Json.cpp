// 這是主要 DLL 檔案。

#include "json/json.h"
#include "json/value.h"
#include "libxl.h"
#include <iostream>
#include "libExcel2Json.h"
using namespace std;
using namespace libxl;
using namespace Json;
using namespace libExcel2Json;

wstring StringToWstring(const string str)
{
	unsigned len = str.size() * 2;
	setlocale(LC_CTYPE, "");    
	wchar_t *p = new wchar_t[len];
	mbstowcs(p,str.c_str(),len);
	wstring str1(p);
	delete[] p;
	return str1;
}
string WstringToString(const wstring str)
{
	unsigned len = str.size() * 4;
	setlocale(LC_CTYPE, "");
	char *p = new char[len];
	wcstombs(p,str.c_str(),len);
	string str1(p);
	delete[] p;
	return str1;
}
/*
* ===  PUBLIC FUNCTION  ======================================================================
*         Name:  Excel2Json
*  Description:
* =====================================================================================
*/
Excel2Json::Excel2Json()
{
	_book = nullptr;
	_row =  nullptr;
	_rows =  nullptr;
}

/*
* ===  PUBLIC FUNCTION  ======================================================================
*         Name:  ~Excel2Json
*  Description:
* =====================================================================================
*/
Excel2Json::~Excel2Json()
{
	deleteXlsBook();
	if(_rows)
		delete _rows;
}

/*
* ===  PUBLIC FUNCTION  ======================================================================
*         Name:  toJSON
*  Description:
* =====================================================================================
*/
string Excel2Json::toJSON(const wchar_t * result_path,unsigned int resultFile_type)
{
	string json;
	
	return json;
}

/*
* ===  PUBLIC FUNCTION  ======================================================================
*         Name:  toJSON
*  Description:
* =====================================================================================
*/
string Excel2Json::toJSON(const wchar_t * file_path)
{
	string json;
	Json::FastWriter package;
	_book = readFile( file_path );
	json = package.write(toJSONValue(_book)); 
	return json;
}
/*
* ===  PUBLIC FUNCTION  ======================================================================
*         Name:  toJSON
*  Description:
* =====================================================================================
*/
string Excel2Json::toJSON(const wchar_t * file_path,const wchar_t * result_path,unsigned int resultFile_type)
{
	string json;
	Json::FastWriter package;
	_book = readFile( file_path );
	json = package.write(toJSONValue(_book)); 
	return json;
}
/*
 * ===  PRIVATE FUNCTION  ======================================================================
 *         Name:  deleteXlsBook
 *  Description:
 * =====================================================================================
 */

bool Excel2Json::deleteXlsBook(){
	if(_book){
		_book->release();
		return _book == NULL;
	}
	return true;
}

/*
* ===  PRIVATE FUNCTION  ======================================================================
*         Name:  readFile
*  Description:
* =====================================================================================
*/

Book* Excel2Json::readFile(const wchar_t * file_path){
	Book* book = xlCreateBook();
	if(book)
	{
		if(book->load(file_path))
		{
			if(deleteXlsBook()){
				_book = book;
				return book;
			}
			book->release();
		}
	}
	return book;
}
/*
* ===  PRIVATE FUNCTION  ======================================================================
*         Name:  convertToJson
*  Description:
* =====================================================================================
*/

string Excel2Json::convertToJson(unsigned int resultFile_type){
	string json;
	switch(resultFile_type)
	{
	case RESULT_FILE_TYPE_JSON:
		break;
	case RESULT_FILE_TYPE_EXECL:
		break;
	case RESULT_FILE_TYPE_INI:
		break;
	case NO_RESULT:
		break;
	}
	return json;
}
/*
* ===  PRIVATE FUNCTION  ======================================================================
*         Name:  toJSONValue
*  Description:
* =====================================================================================
*/

Value Excel2Json::toJSONValue(Book * book){
	Value values = Value();
	string * var ;
	if(book)
	{
		XLSRows * rowList = getRows();
		for ( XLSRows::iterator it = rowList->begin() ; it != rowList->end() ; it++ )
		{
			Value value = Value();
			for ( XLSRow::iterator _it = it->begin() ; _it != it->end() ; _it++ ){
				value.append( *_it );
			}
			values.append( value );
		}
	}
	return values;
}
/*
* ===  PRIVATE FUNCTION  ======================================================================
*         Name:  getRow
*  Description:
* =====================================================================================
*/
XLSRow * Excel2Json::getRow(unsigned int index)
{
	XLSRow * value = NULL;
	if(_book)
	{
		Sheet* sheet = _book->getSheet(0);
		if(sheet)
		{
			const wchar_t * _sheet = NULL;
			wstring ID;
			wstring Case;
			int i = 0;
			_sheet = sheet->readStr(index, i++);
			if(_sheet){
				value = new XLSRow();
				do{
					value->push_back(WstringToString(wstring(_sheet)));
					_sheet = sheet->readStr(index, i++);
				}while(_sheet);
			}
		}
	}
	return _row = value;
}
/*
* ===  PRIVATE FUNCTION  ======================================================================
*         Name:  getRows
*  Description:
* =====================================================================================
*/
XLSRows * Excel2Json::getRows()
{
	XLSRows * rows = NULL;
	if(_rows)
		delete _rows;
	if(_book)
	{
		int i = 0;	
		rows = new XLSRows();
		XLSRow * row ;
		do{
			row = getRow(i++);
			if(row != nullptr)
				rows->push_back(*row);
		}while(row != nullptr);
	}
	return _rows = rows;
}
/*
* ===  PRIVATE FUNCTION  ======================================================================
*         Name:  getRows
*  Description:
* =====================================================================================
*/
XLSRows * Excel2Json::getRows(unsigned int start,unsigned int range)
{
	XLSRows * rows = NULL;
	if(_rows)
		delete _rows;
	if(_book)
	{
		int i = start;	
		rows = new XLSRows();
		XLSRow * row ;
		do{
			XLSRow * row = getRow(i++);
			rows->push_back(*row);
		}while(row && i < range);
	}
	return _rows = rows;
}
