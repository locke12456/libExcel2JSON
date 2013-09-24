// libExcel2Json.h

#pragma once
#include "json/json.h"
#include "json/value.h"
#include "libxl.h"
#include <iostream>
#include <list>
using namespace std;
using namespace libxl;
using namespace Json;
#define NO_RESULT				NULL
#define RESULT_FILE_TYPE_JSON	1
#define RESULT_FILE_TYPE_EXECL	2
#define RESULT_FILE_TYPE_INI	3

#define DEFAULAT_ID		"ID"
#define DEFAULAT_CASE	"Case"

string test( const wchar_t * file_path );
wstring StringToWstring(const string str);
string WstringToString(const wstring str);
namespace libExcel2Json {
	typedef list<string> XLSRow;
	typedef list<XLSRow> XLSRows;
	class Excel2Json
	{
	public :
		Excel2Json();
		Excel2Json(const wchar_t * file_path,const wchar_t * result_path);
		~Excel2Json();
		string toJSON(const wchar_t * file_path);
		string toJSON(const wchar_t * result_path,unsigned int resultFile_type);
		string toJSON(const wchar_t * file_path,const wchar_t * result_path,unsigned int resultFile_type);
	private:
		Book* _book;
		
		XLSRow * _row;

		XLSRows * _rows;

		Book* readFile(const wchar_t * file_path);

		string convertToJson(unsigned int resultFile_type = NULL);

		Value toJSONValue(Book * book);

		XLSRow * getRow(unsigned int index);

		XLSRows * getRows();

		XLSRows * getRows(unsigned int start,unsigned int range);

		bool deleteXlsBook();
		
	};
}
