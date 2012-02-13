#include "MyTGL2/DataType/MytComplex.hpp"
#include <stdio.h>

void Print(const Myt::Complex<double>& c){
	printf("%g+%gi\n",c.Re,c.Im);
}

int main(){
	/*
	4.25+2.71i
	4.25+2.71i
	4.25+2.71i
	8.42+9.62i
	8.42+9.62i
	2.03+2.71i
	2.03+2.71i
	-2.03+-2.71i
	-2.14+-4.2i
	-2.14+-4.2i
	3.4854+3.0081i
	3.4854+3.0081i
	3.4854+3.0081i
	-2.1469+36.0062i
	-2.1469+36.0062i
	2.82883+2.44144i
	2.82883+2.44144i
	0.202596+-0.174852i
	0.466838+-0.0976986i
	0.466838+-0.0976986i
	*/
	Myt::Complex<double> c1={3.14,2.71},c2={5.28,6.91},c3;
	c3=c1+1.11;
	Print(c3);
	c3=c1;
	c3+=1.11;
	Print(c3);
	c3=1.11+c1;
	Print(c3);
	c3=c1+c2;
	Print(c3);
	c3=c1;
	c3+=c2;
	Print(c3);
	///
	c3=c1-1.11;
	Print(c3);
	c3=c1;
	c3-=1.11;
	Print(c3);
	c3=1.11-c1;
	Print(c3);
	c3=c1-c2;
	Print(c3);
	c3=c1;
	c3-=c2;
	Print(c3);
	///
	c3=c1*1.11;
	Print(c3);
	c3=c1;
	c3*=1.11;
	Print(c3);
	c3=1.11*c1;
	Print(c3);
	c3=c1*c2;
	Print(c3);
	c3=c1;
	c3*=c2;
	Print(c3);
	///
	c3=c1/1.11;
	Print(c3);
	c3=c1;
	c3/=1.11;
	Print(c3);
	c3=1.11/c1;
	Print(c3);
	c3=c1/c2;
	Print(c3);
	c3=c1;
	c3/=c2;
	Print(c3);
	///
	return 0;
}
