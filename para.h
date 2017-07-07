#ifndef _PARA_H_
#define _PARA_H_



class Para
{
public:
	int			m_nLenth;
	double		m_fBestPrice;	//最优价格
	int			m_nBPpos;		//最优价格位置
	double		m_nArray[3][10];	// 阶梯量 阶梯价 阶梯量差
	
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
			m_fBestPrice=0;	//最优价格
			m_nBPpos=0;		//最优价格位置
	}
	
};
#endif