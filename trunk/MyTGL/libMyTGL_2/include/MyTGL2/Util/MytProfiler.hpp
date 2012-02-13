#ifndef MYTPROFILER_HPP
#define MYTPROFILER_HPP

namespace Myt{
	class Timing{
	private:
		char InternalData[47];
		bool run;
	public:
		Timing();
		//clear the total time.
		void Clear();
		//start timing. call GetMs() after calling this function.
		void Start();
		//stop timing and add the time duration to total time.
		void Stop();
		//get total time (ms).
		double GetMs();
	};

	//get memory usage (MB).
	double MemoryUsage();

	class CPUUsage{
	private:
		char InternalData[32];
	public:
		CPUUsage();
		//get average CPU usage from last calling this function
		//(or the creation of this class) to now.
		//return value is between 0 and 100. (maybe bigger than 100)
		double Get();
	};
}

#endif