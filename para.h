#ifndef _PARA_H_
#define _PARA_H_



class Para
{
public:
	int			m_nLenth;
	double		m_fBestPrice;	//���ż۸�
	int			m_nBPpos;		//���ż۸�λ��
	double		m_nArray[3][10];	// ������ ���ݼ� ��������
	
public:
	void GetBestPrice(int m_nTotal)
	{
		for (int i=m_nLenth-1;i>=0;i--)
		{
			if (m_nTotal>m_nArray[0][i])
			{
				m_fBestPrice=m_nArray[1][i];
				m_nBPpos=i;
				break;
			}
		}
	}

	void init()
	{
		for (int i=0;i<3;i++)
		{
			for (int j=0;j<10;j++)
			{
				m_nArray[i][j]=0;
			}
		}
			m_nLenth=0;
			m_fBestPrice=0;	//���ż۸�
			m_nBPpos=0;		//���ż۸�λ��
	}
	
};
#endif