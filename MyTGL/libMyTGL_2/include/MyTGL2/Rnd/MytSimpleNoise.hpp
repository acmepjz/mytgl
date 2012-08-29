#ifndef MYTSIMPLENOISE_HPP
#define MYTSIMPLENOISE_HPP

namespace Myt{
	class SimpleNoise{
		static const int N0=19990303;
		static const int N1=1515681721;
		static const int N2=-1910728401;
		static const int N3=-1756408387;
		static const int N4=-703619785;
	public:
		static int Noise1(int x,int Seed){
			int a=N0+Seed*N1+x*N2;
			a^=(6*a*a)^((1003*a)>>7);
			return int((((unsigned int)a)>>7)|(((unsigned int)a)<<25));
		}
		static int Noise2(int x,int y,int Seed){
			int a=N0+Seed*N1+x*N2+y*N3;
			a^=(6*a*a)^((1003*a)>>7);
			return int((((unsigned int)a)>>7)|(((unsigned int)a)<<25));
		}
		static int Noise3(int x,int y,int z,int Seed){
			int a=N0+Seed*N1+x*N2+y*N3+z*N4;
			a^=(6*a*a)^((1003*a)>>7);
			return int((((unsigned int)a)>>7)|(((unsigned int)a)<<25));
		}
	};
}

#endif
