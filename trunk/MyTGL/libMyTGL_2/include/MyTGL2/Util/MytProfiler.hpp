#ifndef MYTPROFILER_HPP
#define MYTPROFILER_HPP

namespace Myt{
	/** \brief Timing class.
	\ingroup Utility
	*/
	class Timing{
	private:
		char InternalData[47];
		bool run;
	public:
		/** Constructor. */
		Timing();
		/** clear the total time. */
		void Clear();
		/** start timing. call GetMs() after calling this function. */
		void Start();
		/** stop timing and add the time duration to total time. */
		void Stop();
		/** get total time (ms). */
		double GetMs();
	};

	/** \brief Get memory usage (MB).
	\ingroup Utility
	*/
	double MemoryUsage();

	/** \brief Get average CPU usage.
	\ingroup Utility
	*/
	class CPUUsage{
	private:
		char InternalData[32];
	public:
		/** Constructor. */
		CPUUsage();
		/// Get average CPU usage from last calling this function
		/// (or the creation of this class) to now.
		/// \return percentage between 0 and 100. (maybe bigger than 100)
		double Get();
	};
}

#endif