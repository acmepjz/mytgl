#include "MyTGL2/Util/MytProfiler.hpp"
#include "MyTGL2/Util/MytStaticAssert.hpp"

#ifdef WIN32

#include <windows.h>
#include <psapi.h>
#pragma comment(lib,"psapi.lib")

struct TimingInternalData{
	LARGE_INTEGER t1,t2,freq;
};

struct CPUUsageInternalData{
	LARGE_INTEGER t_cpu,t,freq;
};

#else

#include <sys/time.h>
#include <sys/resource.h>
#include <stdint.h>
#include <stdlib.h>

struct TimingInternalData{
	double t1;
	timeval t2;
};

struct CPUUsageInternalData{
	timeval t_cpu;
	timeval t;
};

#endif

static const Myt::StaticAssert<sizeof(TimingInternalData)<=sizeof(Myt::Timing)-1> dummy1;
static const Myt::StaticAssert<sizeof(CPUUsageInternalData)<=sizeof(Myt::CPUUsage)> dummy2;

namespace Myt{
	Timing::Timing(){
		TimingInternalData *t=(TimingInternalData*)InternalData;
#ifdef WIN32
		t->t1.QuadPart=0;
		QueryPerformanceFrequency(&t->freq);
#else
		t->t1=0.0;
#endif
		run=false;
	}

	void Timing::Clear(){
		TimingInternalData *t=(TimingInternalData*)InternalData;
#ifdef WIN32
		t->t1.QuadPart=0;
#else
		t->t1=0.0;
#endif
		run=false;
	}

	void Timing::Start(){
		if(!run){
			TimingInternalData *t=(TimingInternalData*)InternalData;
#ifdef WIN32
			QueryPerformanceCounter(&t->t2);
#else
			gettimeofday(&t->t2,NULL);
#endif
			run=true;
		}
	}

	void Timing::Stop(){
		if(run){
			TimingInternalData *t=(TimingInternalData*)InternalData;
#ifdef WIN32
			LARGE_INTEGER a;
			QueryPerformanceCounter(&a);
			t->t1.QuadPart+=(a.QuadPart-(t->t2.QuadPart));
#else
			timeval a;
			gettimeofday(&a,NULL);
			t->t1+=double(a.tv_sec-(t->t2.tv_sec))*1000.0
				+double(a.tv_usec-(t->t2.tv_usec))/1000.0;
#endif
			run=false;
		}
	}

	double Timing::GetMs(){
		TimingInternalData *t=(TimingInternalData*)InternalData;
		if(run){
#ifdef WIN32
			LARGE_INTEGER a;
			QueryPerformanceCounter(&a);
			return double(t->t1.QuadPart+a.QuadPart-(t->t2.QuadPart))/double(t->freq.QuadPart)*1000.0;
#else
			timeval a;
			gettimeofday(&a,NULL);
			return t->t1+double(a.tv_sec-(t->t2.tv_sec))*1000.0
				+double(a.tv_usec-(t->t2.tv_usec))/1000.0;
#endif
		}else{
#ifdef WIN32
			return double(t->t1.QuadPart)/double(t->freq.QuadPart)*1000.0;
#else
			return t->t1;
#endif
		}
	}

	double MemoryUsage(){
#ifdef WIN32
		PROCESS_MEMORY_COUNTERS t;
		t.cb=sizeof(t);
		GetProcessMemoryInfo(GetCurrentProcess(),&t,sizeof(t));
		return (double)(t.WorkingSetSize)/1048576.0;
#else
		//TODO:
		return 0.0;
#endif
	}

	CPUUsage::CPUUsage(){
		CPUUsageInternalData *t=(CPUUsageInternalData*)InternalData;
#ifdef WIN32
		LARGE_INTEGER t1,t2,t3,t4;
		GetProcessTimes(GetCurrentProcess(),(LPFILETIME)&t1,(LPFILETIME)&t2,(LPFILETIME)&t3,(LPFILETIME)&t4);
		QueryPerformanceCounter(&t->t);
		QueryPerformanceFrequency(&t->freq);
		t->t_cpu.QuadPart=t3.QuadPart+t4.QuadPart;
#else
		rusage usage;
		gettimeofday(&t->t,NULL);
		getrusage(RUSAGE_SELF,&usage);
		t->t_cpu.tv_sec=usage.ru_utime.tv_sec+usage.ru_stime.tv_sec;
		t->t_cpu.tv_usec=usage.ru_utime.tv_usec+usage.ru_stime.tv_usec;
#endif
	}

	double CPUUsage::Get(){
		CPUUsageInternalData *t=(CPUUsageInternalData*)InternalData;
#ifdef WIN32
		LARGE_INTEGER t1,t2,t3,t4;

		//get value
		GetProcessTimes(GetCurrentProcess(),(LPFILETIME)&t1,(LPFILETIME)&t2,(LPFILETIME)&t3,(LPFILETIME)&t4);
		QueryPerformanceCounter(&t1);
		t3.QuadPart+=t4.QuadPart;

		//get delta
		t2.QuadPart=t1.QuadPart-t->t.QuadPart;
		t4.QuadPart=t3.QuadPart-t->t_cpu.QuadPart;

		//update
		t->t.QuadPart=t1.QuadPart;
		t->t_cpu.QuadPart=t3.QuadPart;

		//over
		if(t2.QuadPart>0){
			return double(t4.QuadPart)*double(t->freq.QuadPart)/double(t2.QuadPart)/100000.0;
		}else{
			return 0.0;
		}
#else
		rusage usage;
		timeval t0;

		//get value
		gettimeofday(&t0,NULL);
		getrusage(RUSAGE_SELF,&usage);
		usage.ru_utime.tv_sec+=usage.ru_stime.tv_sec;
		usage.ru_utime.tv_usec+=usage.ru_stime.tv_usec;

		//get delta
		double dt=double(t0.tv_sec-t->t.tv_sec)
			+double(t0.tv_usec-t->t.tv_usec)/1000000.0;
		double dt_cpu=double(usage.ru_utime.tv_sec-t->t_cpu.tv_sec)*100.0
			+double(usage.ru_utime.tv_usec-t->t_cpu.tv_usec)/10000.0;

		//update
		t->t=t0;
		t->t_cpu=usage.ru_utime;

		//over
		if(dt>0.000001){
			return dt_cpu/dt;
		}else{
			return 0.0;
		}
#endif
	}
}
