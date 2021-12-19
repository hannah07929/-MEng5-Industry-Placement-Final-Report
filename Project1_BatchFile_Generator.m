%code for exp1 and exp2, when axis in middle

clc; clear all; close all;
    
pitch= 100;  %distance between the backs of two blinds (MEASURE)
Dst_FB=105;%distance from back of blind to the front of the blind (MEASURE)
Dst_RL=2250;%distance between the two sides of the blind(MEASURE)
Dst_cols=[20, 60, 100, 140, Dst_RL/2];%distance between each collumn of sensors and right side of the blind (MEASURE)

Dst_Rws=28; %distance between rows

Dst_ax=Dst_FB/2;%Distance between ax and front of frame also the distance from front of blind to the axis, parallel to slat

Dst_rows=[Dst_ax-Dst_Rws, Dst_ax, Dst_ax+Dst_Rws];   %distance between each row of sensors and front of the blind

Dst_rows_ax = [Dst_ax-Dst_rows(1),Dst_ax-Dst_rows(2),Dst_ax-Dst_rows(3)];%parallel distance between each row of sensors and the axis



%Azimuth
a_z = -180:180; %Field of view. Directly at front is 0°
a_z = -180:5:180; %Field of view. Directly at front is 0°



%%
%Let user chose which exp they want to run:
user_input = menu('Please select 1 for EXP1 and 2 for EXP2.','exp1','exp2'); % ask

switch user_input
    case 1
        disp('running exp1')
        tilt = [20,60,85]; %How much have the blindes opened. Fully closed is 90°, fully open is 0°
        
        %Save the hor files into the Experiment1 folder
        hor_path= sprintf('\\\\data-be\\data-ti-2019\\eit\\50_Labore\\T016-Photovoltaik_1\\06_Projekte\\02_Aktiv\\2019_Schenker_Storen\\DOCS_Hannah\\Case2 (Axis in Middle)\\02_Generate_Batch_File\\02-Case2_Batch_Files\\Case2_EXP1\\');
       
    case 2
        disp('running exp2')
        tilt = [0:10:80 85]; %How much have the blindes opened. Fully closed is 90°, fully open is 0°
        Dst_cols=Dst_RL/2;%distance between each collumn of sensors and right side of the blind (MEASURE)
        
        %Save the hor files into the Experiment2 folder
        hor_path= sprintf('\\\\data-be\\data-ti-2019\\eit\\50_Labore\\T016-Photovoltaik_1\\06_Projekte\\02_Aktiv\\2019_Schenker_Storen\\DOCS_Hannah\\Case2 (Axis in Middle)\\02_Generate_Batch_File\\02-Case2_Batch_Files\\Case2_EXP2\\');
                
end


%%


%Calculations for side shading:


for t=1:length(tilt)       %loop 3 times for the 3 tilting possibilities
    for r = 1:length(Dst_rows) %loop 3 times for the 3 rows
        
        d1=Dst_rows(r)*cos(deg2rad(tilt(t))); %d1=Dst_rows*cos(tilt), where d1 is the adjacent of the horizontal line triangle regarding self shading
        
        %___________________________________debug d1________________________________
        %fprintf('tilt=%.0f, row=%.1f, d1=%.2f\n', tilt(t),Dst_rows(r),d1) %to debug here
        %disp([tilt(t),Dst_rows(r),d1]) %to debug in excel
        % disp(d1)
        %___________________________________________________________________________
        
        d2=Dst_ax-(Dst_rows_ax(r)*cos(deg2rad(tilt(t))));
        %___________________________________debug d1________________________________
        %fprintf('tilt=%.0f, Dst_rows_ax(r)=%.2f, d2=%.2f\n', tilt(t),rd,d2) %to debug here
        %disp([tilt(t),Dst_rows_ax(r),d2]) %to debug in excel
        %disp(d2)
        %___________________________________________________________________________
        
        
        h=pitch-Dst_ax*sin(deg2rad(tilt(t)))+sin(deg2rad(tilt(t)))*Dst_rows_ax(r); %vertical horizon line
        %___________________________________debug d1________________________________
        %fprintf('tilt=%.0f, Dst_rows_ax=%.2f, h=%.2f\n', tilt(t), Dst_rows_ax(r), h) %to debug here
        %disp([tilt(t),Dst_rows_ax(r),h]) %to debug in excel
        % disp(h)
        %___________________________________________________________________________
        
        for c=1:length(Dst_cols)%loop 5 times for the 5 columns
            
            alpha_S=hor(h, d1, d2, Dst_cols(c), a_z); %call function
            alpha_E=[alpha_S(90/(360/(length(alpha_S)-1)):end) alpha_S(1:90/(360/(length(alpha_S)-1))-1)]; %shift alpha_S by 90 deg
            
            hor_Mxt(t,r,c).alphaS = alpha_S; %alpha values for south
            hor_Mxt(t,r,c).alphaE = alpha_E; %alpha values for east i.e. same as alpha values for south but shifted 90 deg
            
            zeroHor = zeros(size(alpha_S));  %alpha values for the zero hor file i.e. all alpha values=0
            
        end
    end
        %___________________________________debug things________________________________
        %figure()
        %plot(alpha_S)
        %ylim([0 90])
        %hold on
        %___________________________________________________________________________
end
%___________________________________debug alpha_S and alpha_E________________________________
%              plot(a_z,alpha_S,'b'); %blue curve is south
%              xlabel('az'); ylabel('alpha');
%              hold on
%              plot(a_z,alpha_E,'r');%red curve is east
%%
%File Generating:

%Generate 91 txt files

%zero hor file
horFile_zeros = [a_z' zeroHor'];            %content for zero hor file

horFileNameZero = ('horZero.hor');
%zerohor_Fullpath=append(hor_path,horFileNameZero);
zerohor_Fullpath=strcat(hor_path,horFileNameZero);
csvwrite(zerohor_Fullpath, horFile_zeros);      %save content for zero hor with horZero label

%Generate horizon files
%for o=1:length(orientation)
    for t=1:length(tilt)
        for r=1:length(Dst_rows)
            for c=1:length(Dst_cols)
                
                horFile_AlphaS = [a_z' round(hor_Mxt(t,r,c).alphaS')];    %content for South hor files
                horFile_AlphaE = [a_z' round(hor_Mxt(t,r,c).alphaE')];    %content for East hor files
                
                horFileS = sprintf('t%d_r%d_c%d_S.hor',t,r,c);
                horFileE = sprintf('t%d_r%d_c%d_E.hor',t,r,c);
                
                %horFileNameS = append(hor_path,horFileS);
                %horFileNameE = append(hor_path,horFileE);
                horFileNameS = strcat(hor_path,horFileS);
                horFileNameE = strcat(hor_path,horFileE);
                
                csvwrite(horFileNameS,horFile_AlphaS);      %save content for South with south label
                csvwrite(horFileNameE,horFile_AlphaE);      %save content for East with East label
                
            end
        end
    end
%end



%%
%Generate excel batch file


%name   longitude  altitude    path             azimuth     inclination
%File1  constant   something   C:\\anypath1     ?            constant
%File2  constant   something   C:\\anypath2     ?            constant
%...
%File91 constant   something   C:\\anypath91    ?            constant

orientation = [0, -90];    %South, East  orientation
latitude    =  47.055;       %Â°E Lat (obtained from Meteonorm)
longitude   =  7.624;       %Â°E Lon (obtained from Meteonorm)
altitude    = 547;          %in meters (obtained from Meteonorm)
inclination = 0;           %what is the inclination referring to? is it not the same as tilt? because the azimuth vs. alpha files already contain the tilt information.

BatchFileName = 'Batch.csv';
%BatchName = append(hor_path,BatchFileName);
BatchName = strcat(hor_path,BatchFileName);
fileID = fopen(BatchName,'w'); %open file for writing; discard existing contents
fprintf(fileID,'Formattyp = 1\n');
fprintf(fileID,'Save = month\n');
fprintf(fileID,'name ; longitude ; latitude ; altitude ; hor ;  azimuth ; inclination \n'); %column titles

for o=1:length(orientation) %loop 2 times for south, west and east
    for t=1:length(tilt)       %loop 10 times for the 10 tilting possibilities
        for r = 1:length(Dst_rows) %loop 3 times for the 3 rows
            for c=1:length(Dst_cols)
                
                if o==1
                    south_files=sprintf('t%d_r%d_c%d_S.hor',t,r,c);
                    %filePath = append(hor_path, south_files);
                    filePath = strcat(hor_path, south_files);
                    simName  = sprintf('t%d_r%d_c%d_S',t,r,c);
                elseif o==2
                    east_files=sprintf('t%d_r%d_c%d_E.hor',t,r,c);
                    %filePath = append(hor_path, east_files);
                    filePath = strcat(hor_path, east_files);
                    simName  = sprintf('t%d_r%d_c%d_E',t,r,c);
                end
                
                fprintf(fileID, strcat(simName, ' ;'));             %write file name
                fprintf(fileID,'%d ;', longitude);                  %write longitude and altitude
                fprintf(fileID,'%d ;', latitude);                   %write longitude and altitude
                fprintf(fileID,'%d ;', altitude);                   %write longitude and altitude
                fprintf(fileID,'%s ;', filePath);                   %write path
                fprintf(fileID,'%d ;', orientation(o));             %write orientation
                fprintf(fileID,'%d ;', tilt(t));                    %write inclination
                fprintf(fileID,'\n');                               %next line
            end
        end
        
        %for the zero file. in the batch file, the zero files depend on
        %tilt and orienation 
        simName = sprintf('ZeroFile_o%d_t%d',o,t);
        %filePath = append(hor_path, horFileNameZero );
        filePath = strcat(hor_path, horFileNameZero );
        fprintf(fileID, strcat(simName, ' ;'));             %write file name
        fprintf(fileID,'%d ;', longitude);                  %write longitude and altitude
        fprintf(fileID,'%d ;', latitude);                   %write longitude and altitude
        fprintf(fileID,'%d ;', altitude);                   %write longitude and altitude
        fprintf(fileID,'%s ;', filePath);                   %write path
        fprintf(fileID,'%d ;', orientation(o));             %write orientation
        fprintf(fileID,'%d ;', tilt(t));                    %write inclination
        fprintf(fileID,'\n');
    end
end
fclose(fileID);

%%
%{
alpha function:
This function outputs the angle alpha for a given height, d1, d2, collumn and azimuth
%}

function alpha=hor(h, d1, d2, c, a_z)
alpha=rad2deg(atan(cosd(a_z)*h/d1));
alpha(alpha<0)=0;
az_frame=rad2deg(atan(c/d2));
alpha(a_z>az_frame)=0;
end


