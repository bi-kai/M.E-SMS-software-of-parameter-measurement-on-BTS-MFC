# ��վ����ƴ��˵��
cid + bsic + cidNum(���߱��5λ) + back�����ƽ�� + downNum�����������4λ�� + fre + front��ǰ��ƽ�� + heading + high�����θ߶ȣ�
+ lac + latitude + level����ƽ�ȣ� + longitude + pitch + roll + rxl���ź�ǿ�ȣ�Ŀ�����ֵ��Ĳ�׼�� + ratio

cid + bsic + cidNum(���߱��5λ) + ����ƽ +���������+Ƶ�� +ǰ���ƽ +����� +���θ߶�+ lac + ά�� + level����ƽ�ȣ� +���� +��� + ����� + rxl���ź�ǿ�ȣ�Ŀ�����ֵ��Ĳ�׼�� + ratio

# data.txt
�����Щ��Ϣ���Ƕ������ļ�����data.txt�С�
	������ռ�Ĵ�С��
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
