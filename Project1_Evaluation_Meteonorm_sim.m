clc; clear all; close all; fclose all;


%%
try
    % See if there is an existing instance of Excel running.
    % If Excel is NOT running, this will throw an error and send us to the catch block below.
    Excel = actxGetRunningServer('Excel.Application');
    Excel.DisplayAlerts = false;
    % If there was no error, then we were able to connect to it.
    Excel.Quit; % Shut down Excel.
catch
end

%%
months=1:12;
orientation = [0, -90];
row = 1:3;
col = 1:5;

%matrices
ResultMtx       =   [];
ResultMtxPrct   =   [];
ResultMtxMonth  =   [];
ResultMtxMonth  = ones(3962,5)*nan;
Zeros_irr = [];

%%
%Let user chose which exp they want to run:
user_input = menu('Please select which experiment you wish to evaluate.','exp1','exp2'); % pop-up menu

switch user_input
    case 1
        disp('evaluating exp1...')
        tilt = [20,60,85]; %How much have the blindes opened. Fully closed is 90°, fully open is 0°
        
        %get Meteonorm simulation monthly files from Experiment1 folder
        pathResults = sprintf('\\\\data-be\\data-ti-2019\\eit\\50_Labore\\T016-Photovoltaik_1\\06_Projekte\\02_Aktiv\\2019_Schenker_Storen\\DOCS_Hannah\\Case2 (Axis in Middle)\\03_Meteonorm_Output_Files\\Case2_EXP1_Monthly\\');
        Result_File='Case2_EXP1_Results.xlsx'; %save result for exp1 into this excel file
        range='A5';
        
    case 2    
        disp('evaluating exp2...')
        tilt = [0:10:80 85]; %How much have the blindes opened. Fully closed is 90°, fully open is 0°
        col=col(1);%only need the central column for exp2
        
        %get Meteonorm simulation monthly files from Experiment2 folder
        pathResults = sprintf('\\\\data-be\\data-ti-2019\\eit\\50_Labore\\T016-Photovoltaik_1\\06_Projekte\\02_Aktiv\\2019_Schenker_Storen\\DOCS_Hannah\\Case2 (Axis in Middle)\\03_Meteonorm_Output_Files\\Case2_EXP2_Monthly\\');
        Result_File='Case2_EXP2_Results.xlsx'; %save result for exp2 into this excel file
        range='B2';
        
    case 0
        disp('Program ended.')
        return
end


%%
%Read files from meteonorm and calculate the values for the Res_irr for all situations

for o=1:length(orientation) %loop 2 times for south and east
    for t=1:length(tilt)       %loop 3 or 10 times for the 3 or 10 tilting possibilities, this also represents the rows in the matric of irr
        
        %get zero meteo files to read
        tempFileName2 = sprintf('ZeroFile_o%d_t%d-mon.txt',o,t); %get the zero file
        File_from_Meteo2 = append(pathResults, tempFileName2); %Combine path and file name then assign it to the variable 'File_from_Meteo' to access the file
        
        %define table
        T = readtable(File_from_Meteo2,'Delimiter', '\t', 'Range', '11:24'); %table starts at line 11 ends at line 24
        if tilt(t)==0
            Zeros_irr = T{1:12,'H_Gh'}; %row m of column "H_Gh" of the table T
        else
            Zeros_irr = T{1:12,'H_Gk'}; %row m of column "H_Gk" of the table T
        end
        
        subsetResults = []; %empty out this matrix
        
        for r = 1:length(row)
            for c=1:length(col)
                
                %get non-zero meteo files to read
                if o==1 %if south, get the file with an 'S' in the file name
                    tempFileName = sprintf('t%d_r%d_c%d_S-mon.txt',t,r,c);
                elseif o==2 %if east, get the file with an 'E' in the file name
                    tempFileName = sprintf('t%d_r%d_c%d_E-mon.txt',t,r,c);
                end
                
                File_from_Meteo = append(pathResults, tempFileName); %Combine path and file name
                T = readtable(File_from_Meteo, 'Delimiter', '\t', 'Range', '13:26'); %table starts at line 13
                
                if tilt(t)==0 %if there is no tilt, read colum "H_Ghhor"
                    subset = T{1:12, 'H_Ghhor'}; %row m of column "H_Ghhor" of the table T
                    
                    %Otherewise the txt file represents a situation with a tilt, therefore read colum "H_Gkhor"
                else
                    subset = T{1:12,'H_Gkhor'}; %row m of column "H_Gkhor" of the table T
                    
                end
                
                Results(t,o,r,c).Res_irr = subset; %the value of irr for a given m,o,r,c and t
                Results(t,o,r,c).Diff_irr = minus(Zeros_irr, Results(t,o,r,c).Res_irr); %The actaul irr value
                Results(t,o,r,c).Diff_irr_sum = sum(Results(t,o,r,c).Diff_irr); %sum of all irr values for 12 months
                
                subsetResults(r,c) = Results(t,o,r,c).Diff_irr_sum; %puts actual irr values in a temporary matrix, so it's easier to be dealt with later
                subsetResultsMonth(r,c,1:12) = Results(t,o,r,c).Diff_irr; %put sum of all irr values for 12 months in a temporary matrix, so it's easier to be dealt with later
                

                    if user_input==1
                        for m=1:length(months)
                            ResultMtxMonth(1+(m-1)*25 + (o-1)*15 + (r-1)*5 + (c-1)*305+ (t-1),1) = m;
                            ResultMtxMonth(1+(m-1)*25 + (o-1)*15 + (r-1)*5 + (c-1)*305+ (t-1),2) = orientation(o);
                            ResultMtxMonth(1+(m-1)*25 + (o-1)*15 + (r-1)*5 + (c-1)*305+ (t-1),3) = r;
                            ResultMtxMonth(1+(m-1)*25 + (o-1)*15 + (r-1)*5 + (c-1)*305+ (t-1),4) = c;
                            ResultMtxMonth(1+(m-1)*25 + (o-1)*15 + (r-1)*5 + (c-1)*305+ (t-1),5) = tilt(t);
                            ResultMtxMonth(1+(m-1)*25 + (o-1)*15 + (r-1)*5 + (c-1)*305+ (t-1),6) = Results(t,o,r,c).Diff_irr(m);
                        end

                    elseif user_input==2
                        for m=1:length(months)
                            ResultMtxMonth(1+(m-1)*66 + (o-1)*33 + (r-1)*11 + (t-1),1) = m;
                            ResultMtxMonth(1+(m-1)*66 + (o-1)*33 + (r-1)*11 + (t-1),2) = orientation(o);
                            ResultMtxMonth(1+(m-1)*66 + (o-1)*33 + (r-1)*11 + (t-1),3) = r;
                            ResultMtxMonth(1+(m-1)*66 + (o-1)*33 + (r-1)*11 + (t-1),4) = tilt(t);
                            ResultMtxMonth(1+(m-1)*66 + (o-1)*33 + (r-1)*11 + (t-1),5) = Results(t,o,r,c).Diff_irr(m);
                        end
                    end


            end
            
        end
        
        %annual 
        if user_input==1
            ResultMtx =     [ResultMtx;     [subsetResults, [nan; nan; nan], (subsetResults/max(max(subsetResults))*100), [nan; nan; nan],  [tilt(t); tilt(t); tilt(t)], [orientation(o);  orientation(o); orientation(o)] ]];
            ResultMtx =     [ResultMtx; ones(1,14)*nan];

        elseif user_input==2
           
            if o==1
                ResultMtx(1,:) = [nan tilt]; %the first row of matrix consists of the tilt angles
                ResultMtx(2:5,1) = [1; 2; 3; nan]; %assign values 1,2,3 (for the 3 rows) to the second to fourth element of the first column
                ResultMtx(2:4,t+1) = subsetResults; %put the irr values in the right places in the matrix
                ResultMtx(5,t+1)   = mean(ResultMtx(2:4,t+1)); %average irr of all three rows per angle 
                ResultMtx =     [ResultMtx; ones(1,11)*nan];
            elseif o==2
                ResultMtx(7,:) = [nan tilt];
                ResultMtx(8:11,1) = [1; 2; 3;nan];
                ResultMtx(8:10,t+1) = subsetResults;
                ResultMtx(11,t+1)   = mean(ResultMtx(8:10,t+1));
            end
        end
    end
end


    
    %%

    xlswrite(Result_File,[ResultMtx ResultMtxPrct],'annual', range); %display the irradiance value and the percentage
   % xlswrite(Result_File,[ResultMtxMonth],'monthly', 'A5'); %display the irradiance value and the percentage
    

    xlswrite(Result_File,[ResultMtxMonth ResultMtxPrct],'monthly', 'A2'); %display the irradiance value and the percentage


    winopen(Result_File);
    %%
    
    disp('Done.')
    
    
    
    
    
    
    
    
    
    
