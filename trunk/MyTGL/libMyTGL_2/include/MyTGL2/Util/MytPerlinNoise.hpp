#ifndef MYTPERLINNOISE_HPP
#define MYTPERLINNOISE_HPP

namespace Myt{
	template<class T_NoiseProvider>
	class PerlinNoise{
	public:
		static float Noise2(float x,float y,int Seed,T_NoiseProvider& noise){
			//???
			const float G[16][2]={
				{1,1},{-1,1},{1,-1},{-1,-1},
				{1,0},{-1,0},{1,0},{-1,0},
				{0,1},{0,-1},{0,1},{0,-1},
				{1,1},{-1,1},{0,-1},{0,-1}
			};

			int x0=int(x)-(x>0?0:1);
			int y0=int(y)-(y>0?0:1);
			float dx=x-float(x0);
			float dy=y-float(y0);

			int idx;
			
			idx=noise.Noise2(x0,y0,Seed)&0xF;
			float g00=G[idx][0]*dx+G[idx][1]*dy;
			idx=noise.Noise2(x0,y0+1,Seed)&0xF;
			float g01=G[idx][0]*dx+G[idx][1]*(dy-1.0f);
			idx=noise.Noise2(x0+1,y0,Seed)&0xF;
			float g10=G[idx][0]*(dx-1.0f)+G[idx][1]*dy;
			idx=noise.Noise2(x0+1,y0+1,Seed)&0xF;
			float g11=G[idx][0]*(dx-1.0f)+G[idx][1]*(dy-1.0f);

			dx*=dx*dx*(dx*(6.0f*dx-15.0f)+10.0f);
			dy*=dy*dy*(dy*(6.0f*dy-15.0f)+10.0f);

			g00+=(g01-g00)*dy;
			g10+=(g11-g10)*dy;

			g00+=(g10-g00)*dx;

			return g00;
		}
		static float Noise3(float x,float y,float z,int Seed,T_NoiseProvider& noise){
			const float G[16][3]={
				{1,1,0},{-1,1,0},{1,-1,0},{-1,-1,0},
				{1,0,1},{-1,0,1},{1,0,-1},{-1,0,-1},
				{0,1,1},{0,-1,1},{0,1,-1},{0,-1,-1},
				{1,1,0},{-1,1,0},{0,-1,1},{0,-1,-1}
			};

			int x0=int(x)-(x>0?0:1);
			int y0=int(y)-(y>0?0:1);
			int z0=int(z)-(z>0?0:1);
			float dx=x-float(x0);
			float dy=y-float(y0);
			float dz=z-float(z0);

			int idx;
			
			idx=noise.Noise3(x0,y0,z0,Seed)&0xF;
			float g000=G[idx][0]*dx+G[idx][1]*dy+G[idx][2]*dz;
			idx=noise.Noise3(x0,y0,z0+1,Seed)&0xF;
			float g001=G[idx][0]*dx+G[idx][1]*dy+G[idx][2]*(dz-1.0f);
			idx=noise.Noise3(x0,y0+1,z0,Seed)&0xF;
			float g010=G[idx][0]*dx+G[idx][1]*(dy-1.0f)+G[idx][2]*dz;
			idx=noise.Noise3(x0,y0+1,z0+1,Seed)&0xF;
			float g011=G[idx][0]*dx+G[idx][1]*(dy-1.0f)+G[idx][2]*(dz-1.0f);
			idx=noise.Noise3(x0+1,y0,z0,Seed)&0xF;
			float g100=G[idx][0]*(dx-1.0f)+G[idx][1]*dy+G[idx][2]*dz;
			idx=noise.Noise3(x0+1,y0,z0+1,Seed)&0xF;
			float g101=G[idx][0]*(dx-1.0f)+G[idx][1]*dy+G[idx][2]*(dz-1.0f);
			idx=noise.Noise3(x0+1,y0+1,z0,Seed)&0xF;
			float g110=G[idx][0]*(dx-1.0f)+G[idx][1]*(dy-1.0f)+G[idx][2]*dz;
			idx=noise.Noise3(x0+1,y0+1,z0+1,Seed)&0xF;
			float g111=G[idx][0]*(dx-1.0f)+G[idx][1]*(dy-1.0f)+G[idx][2]*(dz-1.0f);

			dx*=dx*dx*(dx*(6.0f*dx-15.0f)+10.0f);
			dy*=dy*dy*(dy*(6.0f*dy-15.0f)+10.0f);
			dz*=dz*dz*(dz*(6.0f*dz-15.0f)+10.0f);

			g000+=(g001-g000)*dz;
			g010+=(g011-g010)*dz;
			g100+=(g101-g100)*dz;
			g110+=(g111-g110)*dz;

			g000+=(g010-g000)*dy;
			g100+=(g110-g100)*dy;

			g000+=(g100-g000)*dx;

			return g000;
		}
	};
}

#endif

