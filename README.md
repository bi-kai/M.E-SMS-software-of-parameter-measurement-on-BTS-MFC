# 基站参数拼接说明
cid + bsic + cidNum(天线编号5位) + back（后电平） + downNum（电子下倾角4位） + fre + front（前电平） + heading + high（海拔高度）
+ lac + latitude + level（电平比） + longitude + pitch + roll + rxl（信号强度，目测这个值测的不准） + ratio

cid + bsic + cidNum(天线编号5位) + 后瓣电平 +电子下倾角+频段 +前瓣电平 +方向角 +海拔高度+ lac + 维度 + level（电平比） +经度 +倾角 + 横滚角 + rxl（信号强度，目测这个值测的不准） + ratio

# data.txt
存放这些信息的是二进制文件，在data.txt中。
	下面是占的大小：
     memset(info2->mem2_cid,'\0',5);
     memset(info2->mem2_bsic,'\0',3);
     memset(info2->mem2_lac,'\0',3);
     memset(info2->mem2_fre,'\0',10);
     memset(info2->mem2_longitude,'\0',10);
     memset(info2->mem2_latitude,'\0',10);
     memset(info2->mem2_heading,'\0',10);
     memset(info2->mem2_roll,'\0',10);
     memset(info2->mem2_pitch,'\0',10);
     memset(info2->mem2_ratio,'\0',5);
     memset(info2->mem2_high,'\0',5);
     memset(info2->mem2_cidNum,'\0',6);
     memset(info2->mem2_downNum,'\0',4);
     memset(info2->mem2_level,'\0',3);
     memset(info2->mem2_front,'\0',3);
     memset(info2->mem2_back,'\0',3);
     memset(info2->mem2_rxl,'\0',10);
	 
	 strcat(info2->mem2_cid,cid);
     strcat(info2->mem2_bsic,bsic);
     strcat(info2->mem2_lac,lac);
     strcat(info2->mem2_fre,fre);
     strcat(info2->mem2_longitude,longitude);
     strcat(info2->mem2_latitude,latitude);
     strcat(info2->mem2_heading,heading);
     strcat(info2->mem2_roll,roll);
     strcat(info2->mem2_ratio,ratio);
     strcat(info2->mem2_cidNum,cidNum);
     strcat(info2->mem2_downNum,downNum);
     strcat(info2->mem2_back,back_signal);
     strcat(info2->mem2_front,front_signal);
     strcat(info2->mem2_rxl,rxl);
     strcat(info2->mem2_level,siglevel);
     strcat(info2->mem2_pitch,pitch);
