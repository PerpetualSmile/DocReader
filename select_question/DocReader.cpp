// DocReader.cpp : �������̨Ӧ�ó������ڵ㡣
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


// Ψһ��Ӧ�ó������

CWinApp theApp;

using namespace std;

bool compare(int a,int b)
{
	return a < b;
}

void randomRead(string s, vector<string>&result)
{
	string::size_type position,end;
	vector<string> temp;
	map<int, int> nummap;
	vector<int> randomNum;
	stringstream rank;
	int p = 1;
	string q;
	string buff;
	position = s.find("1.");
	s = s.substr(position);
	while (true)
	{
		p++;
		rank.clear();
		rank << p;
		rank >> q;
		end = s.find("[");
		if (end == string::npos)
		{
			temp.push_back(s);
			break;
		}
		buff = s.substr(0, end);
		temp.push_back(buff);
		position = s.find(q + ".");
		s = s.substr(position);
	}
	int num;
	int length = temp.size();
	srand((unsigned)time(0));
	cout <<"�ܹ���ȡ��Ŀ����" <<length << endl;
	int n;
	cout << "��������Ҫ��������";
	cin >> n;
	for(unsigned int i=0;i<n;i++)
	{
		num = rand() % length;
		//cout << "num:" << num << endl;
		if (nummap.find(num)!=nummap.end())
		{
			i--;
			continue;
		}
		nummap[num] = 1;
		randomNum.push_back(num);
	}
	cout <<"���ѡȡ����Ŀ������"<< randomNum.size() << endl<<"��ţ�"<<endl;
	std::sort(randomNum.begin(),randomNum.end(),compare);
	for (unsigned int i=0;i<randomNum.size();i++)
	{
		cout << randomNum[i]+1<<endl;
		result.push_back(temp[randomNum[i]]);
	}

}

int main()
{
    int nRetCode = 0;

    HMODULE hModule = ::GetModuleHandle(nullptr);

    if (hModule != nullptr)
    {
        // ��ʼ�� MFC ����ʧ��ʱ��ʾ����
        if (!AfxWinInit(hModule, nullptr, ::GetCommandLine(), 0))
        {
            // TODO: ���Ĵ�������Է���������Ҫ
            wprintf(L"����: MFC ��ʼ��ʧ��\n");
            nRetCode = 1;
        }
        else
        {
            // TODO: �ڴ˴�ΪӦ�ó������Ϊ��д���롣
			if (CoInitialize(NULL) != S_OK)
			{
				AfxMessageBox(_T("��ʼ��COM֧�ֿ�ʧ��!"));
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
			CDocuments  docs;
			CDocument0  docSource;
			CDocument0  docDestination;
			CRange  aRange;

			COleVariant  vTrue((short)TRUE),
				vFalse((short)FALSE),
				vOpt((long)DISP_E_PARAMNOTFOUND, VT_ERROR);


			if (!(wordApp1.CreateDispatch(_T("word.application"))&&wordApp2.CreateDispatch(_T("word.application")))) //����WORD
			{
				AfxMessageBox(_T("OFFICEû�а�װ?"));
				return 0;
			}


			//wordApp1.put_Visible(true);//����word�Ƿ�ɼ�
			string filepath;
			cout << "���������·����";
			cin >> filepath;
			CString filep;
			filep= filepath.c_str();

			doc1 = wordApp1.get_Documents();
			doc1.AttachDispatch(wordApp1.get_Documents());
			docSource = doc1.Open(COleVariant(filep), vFalse, vTrue, vFalse, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt, vOpt);
			aRange = docSource.Range(vOpt, vOpt);
			CString buff = aRange.get_Text();
			string temp =(CStringA)buff;

			vector<string> result;
			randomRead(temp,result);//�����ȡһ����������Ŀ


			docs = wordApp2.get_Documents();
			docs.AttachDispatch(wordApp2.get_Documents());
			docs.Add(new CComVariant(_T("")), new CComVariant(FALSE), new CComVariant(0), new CComVariant());//�������ĵ�
			CDocument0 doc0 = wordApp2.get_ActiveDocument();
			CSelection select = wordApp2.get_Selection();
			//д���ı�
			CFont0 font = select.get_Font();
			font.put_Name(_T("����"));//��������
			font.put_Size(10);
			font.put_Color(WdColor::wdColorBlack);
			//font.put_Bold(1);

			//select.TypeText(_T("The First Table!"));
			CString x;
			for (unsigned int i=0;i<result.size();i++)
			{
				x = result[i].c_str();
				select.TypeText(x);
			}
			COleVariant vEnd(_T("END")), vT(SHORT(TRUE)), vF(SHORT(FALSE));
			CFile file;
			CString strSaveFile;
			BOOL bUse = TRUE;
			while (bUse)
			{
				CFileDialog fileDialog(FALSE);
				fileDialog.m_ofn.lpstrTitle = _T("����Word�ĵ�");
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
						AfxMessageBox(_T("�ļ����ڱ༭���޷����д洢��"));
					}
				}
				else
				{
					bUse = FALSE;
					AfxMessageBox(_T("ȡ���洢����Ҫʹ�����ݣ����ٴ������"));
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
        // TODO: ���Ĵ�������Է���������Ҫ
        wprintf(L"����: GetModuleHandle ʧ��\n");
        nRetCode = 1;
    }

    return nRetCode;
}

