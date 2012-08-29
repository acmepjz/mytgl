#ifndef MYTLCG_HPP
#define MYTLCG_HPP

namespace Myt{
    //Linear Congruential Generator (LCG) pseudorandom number generator
    //(http://en.wikipedia.org/wiki/Linear_congruential_generator)
    //X_{n+1}=aX_n+c\bmod m
    //The following table lists the parameters of LCGs in common use,
    //including built-in rand() functions in runtime libraries of various compilers.
    /*
    Source                                                          m         a                    c                    output bits of seed in rand() / Random(L)
    Numerical Recipes                                               2^32      1664525              1013904223  
    Borland C/C++                                                   2^32      22695477             1                    bits 30..16 in rand(), 30..0 in lrand()
    glibc (used by GCC)                                             2^32      1103515245           12345                bits 30..0
    ANSI C: Watcom, Digital Mars, CodeWarrior, IBM VisualAge C/C++  2^32      1103515245           12345                bits 30..16
    Borland Delphi, Virtual Pascal                                  2^32      134775813            1                    bits 63..32 of (seed * L)
    Microsoft Visual/Quick C/C++                                    2^32      214013               2531011              bits 30..16
    Microsoft Visual Basic (6 and earlier)                          2^24      1140671485           12820163  
    RtlUniform from Native API                                      2^31 - 1  2147483629           2147483587  
    Apple CarbonLib                                                 2^31 - 1  16807                0                    see MINSTD
    MMIX by Donald Knuth                                            2^64      6364136223846793005  1442695040888963407  
    VAX's MTH$RANDOM, old versions of glibc                         2^32      69069                1  
    Java's java.util.Random                                         2^48      25214903917          11                   bits 47...16
    LC53 in Forth (programming language)                            2^32 - 5  2^32 - 333333333     0  
    */

	//m=2^32
    template<int a=22695477,int c=1,int N_High=30,int N_Low=16>
    class LCG{
	private:
		unsigned int X;
		static const unsigned int Mask=(unsigned int)((1ULL<<(N_High-N_Low+1))-1);
	public:
		void Init(unsigned int s){
			X=s;
		}
		unsigned int Rnd(){
			X=a*X+c;
			return (X>>N_Low)&Mask;
		}
    };
}

#endif
