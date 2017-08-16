// DocReader.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include "DocReader.h"
#include"CDocument0.h"
#include"CApplication.h"
#include"CDocuments.h"
#include"CRange.h"
#include"CFont0.h"
#include"CSelection.h"
#include<iostream>
#include<CString>
#include<string>
#include<fstream>
#include<vector>
#include<sstream>
#include<map>
#include<algorithm>
#ifdef _DEBUG
#define new DEBUG_NEW


#endif


// 唯一的应用程序对象

CWinApp theApp;

using namespace std;

bool compare(int a,int b)
{
	return a < b;
}

void randomRead(string s, map<int, string>&answer, map<int, string>&question)
{
	string::size_type start, end,x;
	vector<int> randomNum;
	int p = 0;
	int q = 0;
	string temp = s;
	string buff;
	while (true)
	{
	    start = s.find("[");
		if (start == string::npos)
		{
			break;
		}
		p++;
	    s = s.substr(start);
		end = s.find("]");
		buff = s.substr(1, end-start-1);
		s = s.substr(end + 1);
		if (p % 5 == 0)
		{
			q++;
			answer[q] = buff;
			//cout << q << " " << answer[q] << endl;
		}
	}
	s = temp;
	p = 0;
	while (true)
	{
		start = s.find("[");
		if (start == string::npos)
		{
			break;
		}
		p++;
		s = s.substr(start);
		end = s.find("[",start+1);
		end= s.find("[",end+1);
		end = s.find("[", end + 1);
		end = s.find("[", end + 1);
		end = s.find("[", end + 1);
		if (end == string::npos)
		{
			buff = s.substr(0);
			question[p] = buff;
			break;
		}
		buff = s.substr(0, end - start);
		s = s.substr(end);
		question[p] = buff;
		//cout << p << endl; //" " << question[p]<< endl;
	}
}
void upperstring(string &buff)
{
	for (unsigned int i = 0; i < buff.size(); i++)
		if (buff[i] >= 'a'&&buff[i] <= 'z')
			buff[i] -= 32;

}
void analyze(string s,map<int,string> &number)
{
	string::size_type start,start1 ,end, x;
	int num;
	int temp;
	int i = 0;
	int num_before=0;
	string buff;
	while (true)
	{
		num = 0;
		temp = 1;
		start = s.find(".");
		start1 = start;
		if (start==string::npos)
		{
			break;
		}
		while (s[start - 1] == 'A'||s[start - 1] == 'B'||s[start - 1] == 'C'||s[start-1] == 'D')
		{
			start = s.find(".",start+1);
			if (start == string::npos)
			{
				break;
			}
			start1 = start;
		}
		if (start == string::npos)
		{
			break;
		}
		while ((start-1)!=string::npos&&s[start-1]>='0'&&s[start-1]<='9')
		{
			num+=((s[start - 1] - '0')*temp);
			temp *= 10;
			start--;
		}
		if (num < num_before) 
		{
			s = s.substr(start1 + 1);
			continue;
		}
		start = s.find("（");
		end = s.find("）",start+1);
		if ((start == string::npos) || (end == string::npos)) break;
		buff=s.substr(start+2,end-start-2);
		upperstring(buff);
		number[num] = buff;
		s = s.substr(end + 1);
		num_before = num;
		//cout << buff<<endl;
	}


}



void pigai(map<int, string>number,map<int,string>answer ,vector<int> &all_question,map<int,int> &wrong_question)
{
	map<int, string>::iterator it;
	for (it=number.begin();it!=number.end();it++)
	{
		all_question.push_back(it->first);
		if (it->second.find(answer[it->first])!=string::npos)continue;
		wrong_question[it->first]=1;
		//cout <<it->first << endl;
	}



}

int main()
{
    int nRetCode = 0;

    HMODULE hModule = ::GetModuleHandle(nullptr);

    if (hModule != nullptr)
    {
        // 初始化 MFC 并在失败时显示错误
        if (!AfxWinInit(hModule, nullptr, ::GetCommandLine(), 0))
        {
            // TODO: 更改错误代码以符合您的需要
            wprintf(L"错误: MFC 初始化失败\n");
            nRetCode = 1;
        }
        else
        {
            // TODO: 在此处为应用程序的行为编写代码。
			if (CoInitialize(NULL) != S_OK)
			{
				AfxMessageBox(_T("初始化COM支持库失败!"));
				return  -1;
			}
			COleVariant varstrNull(_T(""));
			COleVariant varTrue(short(1), VT_BOOL);
			COleVariant varFalse(short(0), VT_BOOL);
			//COleVariant vTure(SHORT(TRUE)), vFalse(SHORT(FALSE));
			COleVariant vE(_T(""));
			COleVariant v0(SHORT(0)), v1(SHORT(1)), v2(SHORT(2)), v5(SHORT(5)), v6(SHORT(6)), v12(SHORT(12)), v22(SHORT(22));
			CString strPath = _T("..\\Perforation.tdt");

			CApplication  wordApp1;
			CApplication  wordApp2;
			CDocuments  doc1;
			CDocuments  doc2;
			CDocuments  docs;
			CDocument0  docSource;
			CDocument0  docDestination;
			CRange  aRange;
			map<int, string> answer;
			map<int, string>question;

			COleVariant  vTrue((short)TRUE),
				vFalse((short)FALSE),
				vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);


			if (!(wordApp1.CreateDispatch(_T("word.application"))&&wordApp2.CreateDispatch(_T("word.application")))) //启动WORD
			{
				AfxMessageBox(_T("OFFICE没有安装?"));
				return 0;
			}


			//wordApp1.put_Visible(true);//设置word是否可见
			string filepath;
			cout << "请输入题库路径：";
			cin >> filepath;
			CString filep;
			filep= filepath.c_str();

			doc1 = wordApp1.get_Documents();
			doc1.AttachDispatch(wordApp1.get_Documents());
			docSource = doc1.Open(COleVariant(filep), vFalse, vTrue, vFalse, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt);
			aRange = docSource.Range(vOpt, vOpt);
			CString buff = aRange.get_Text();
			string temp =(CStringA)buff;
			randomRead(temp,answer,question);//分析题库数据


			cout << "请输入已完成的题目文件路径：";
			cin >> filepath;
			
			filep = filepath.c_str();
			doc2 = wordApp1.get_Documents();
			doc2.AttachDispatch(wordApp1.get_Documents());
			docSource = doc2.Open(COleVariant(filep), vFalse, vTrue, vFalse, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt);
			aRange = docSource.Range(vOpt, vOpt);
			CString buff1 = aRange.get_Text();
			string temp1 = (CStringA)buff1;
			map<int,string> number;
			analyze(temp1,number);//分析结果数据

			vector<int> all_question;
			map<int,int> wrong_question;
			pigai(number,answer,all_question,wrong_question);//批改答题
			
			docs = wordApp2.get_Documents();
			docs.AttachDispatch(wordApp2.get_Documents());
			docs.Add(new CComVariant(_T("")), new CComVariant(FALSE), new CComVariant(0), new CComVariant());//创建新文档
			CDocument0 doc0 = wordApp2.get_ActiveDocument();
			CSelection select = wordApp2.get_Selection();
			//写入文本
			CFont0 font = select.get_Font();
			font.put_Name(_T("宋体"));//设置字体
			font.put_Size(20);
			//font.put_Color(WdColor::wdColorBlack);
			//font.put_Bold(1);

			//select.TypeText(_T("The First Table!"));
			font.put_Color(WdColor::wdColorRed);
			stringstream stream1;
			
			stream1<<all_question.size();
			string string1;
			stream1 >> string1;
			//cout << string1 << endl; 

			stream1.clear();
			stream1 << all_question.size()-wrong_question.size();
			string string2;
			stream1 >> string2;
			//cout << string2 << endl;

			CString x;
			string y = "总题数："+string1;//+all_question.size();
			
			y += "\r";
			string z = "正确题数："+string2;//+ (all_question.size() - wrong_question.size());
			z += "\r";
			x = y.c_str();
			select.TypeText(x);
			x = z.c_str();
			select.TypeText(x);
			font.put_Size(10);
	      	for (unsigned int i=0;i<all_question.size();i++)
			{
				font.put_Color(WdColor::wdColorBlack);
				if (wrong_question.find(all_question[i]) != wrong_question.end())
				{
					font.put_Color(WdColor::wdColorRed);
					question[all_question[i]].insert(question[all_question[i]].find_first_of('（')+1, number[all_question[i]]);
					
				}
				x =question[all_question[i]].c_str();
				select.TypeText(x);
			}
			COleVariant vEnd(_T("END")), vT(SHORT(TRUE)), vF(SHORT(FALSE));
			CFile file;
			CString strSaveFile;
			BOOL bUse = TRUE;
			while (bUse)
			{
				CFileDialog fileDialog(FALSE);
				fileDialog.m_ofn.lpstrTitle = _T("保存Word文档");
				fileDialog.m_ofn.lpstrFilter = _T("Word Document(*.doc)\0*.doc\0All Files(*.*)\0*.*\0\0");
				fileDialog.m_ofn.lpstrDefExt =_T( ".doc");
				if (IDOK == fileDialog.DoModal())
				{
					strSaveFile = fileDialog.GetPathName();
					if (file.Open(strSaveFile, CFile::modeWrite | CFile::modeCreate))
					{
						file.Close();
						bUse = FALSE;
						doc0.SaveAs(COleVariant(strSaveFile), v0, vF, vE, vF, vE, vF, vF, vF, vF, vF, vF, vF, vF, vF, vF);
					}
					else
					{
						bUse = TRUE;
						AfxMessageBox(_T("文件正在编辑，无法进行存储！"));
					}
				}
				else
				{
					bUse = FALSE;
					AfxMessageBox(_T("取消存储！如要使用数据，请再次输出！"));
				}
			}
			font.ReleaseDispatch();
			select.ReleaseDispatch();
			//doc1.Close(vOpt, vOpt, vOpt);
			//docs.Close(vOpt, vOpt, vOpt);
			
			//doc0.ReleaseDispatch();
			//doc1.ReleaseDispatch();
			//docs.ReleaseDispatch();
			wordApp1.Quit(vFalse, vFalse, vFalse);
			wordApp2.Quit(vFalse, vFalse, vFalse);
			//wordApp1.ReleaseDispatch();
			//wordApp2.ReleaseDispatch();
			cout << "Done!" << endl;
			CoUninitialize();

        }
    }
    else
    {
        // TODO: 更改错误代码以符合您的需要
        wprintf(L"错误: GetModuleHandle 失败\n");
        nRetCode = 1;
    }

    return nRetCode;
}

