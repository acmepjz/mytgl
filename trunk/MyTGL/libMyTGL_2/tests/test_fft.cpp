#include "MyTGL2/Util/MytProfiler.hpp"
#include "MyTGL2/Util/MytFFT.hpp"
//#include "MyTGL2/Util/MytFFT_kissfft.hpp"
#include <stdio.h>

//old: Time=524.616ms
//use ShuffleProvider: Time=467.502ms
//new: Time=352.261ms

//kiss_fft Time=427.518ms

int main(){
	//test only
	Myt::Timing t;
	Myt::Complex<double> *d=Myt::Malloc<Myt::Complex<double> >(1024);
	Myt::Complex<double> *d2=Myt::Malloc<Myt::Complex<double> >(1024);
#if 1
	Myt::FFTShuffleProvider sp;
	t.Start();
	sp.Create(1024);
	for(int i=0;i<1024;i++){
		d[i].Re=(double)(i^((i*53)>>5));
		d[i].Im=(double)((i+143)^((i*29)>>5));
	}
	for(int i=0;i<10000;i++){
		Myt::FFT<double,double>::Calc(d,d2,1024,false,true,sp);
	}
	t.Stop();
#else
	Myt::FFTData<double> fftd;
	//===
	t.Start();
	fftd.Create(1024,false);
	for(int i=0;i<1024;i++){
		d[i].Re=(double)(i^((i*53)>>5));
		d[i].Im=(double)((i+143)^((i*29)>>5));
	}
	for(int i=0;i<10000;i++){
		Myt::FFT<double,double>::Calc(d,d2,fftd,1,true);
	}
	t.Stop();
	//===
#endif
	{
		FILE *f=fopen("out.txt","w");
		for(int i=0;i<1024;i++){
			fprintf(f,"%0.6e,%0.6e\n",d2[i].Re,d2[i].Im);
		}
		fclose(f);
	}
	Myt::Free(d);
	Myt::Free(d2);
	//Myt::SaveBmp(bm,"out.bmp");
	printf("Hello, World!\nTime=%0.3fms\n",t.GetMs());
	return 0;
}