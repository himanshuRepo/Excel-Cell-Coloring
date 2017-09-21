%% Cell coloring upto columns defined by two alphabets i.e. 702 (26+676)
clc;
clear all;
close all;
alphabet=['A';
    'B';
    'C';
    'D';
    'E';
    'F';
    'G';
    'H';
    'I';
    'J';
    'K';
    'L';
    'M';
    'N';
    'O';
    'P';
    'Q';
    'R';
    'S';
    'T';
    'U';
    'V';
    'W';
    'X';
    'Y';
    'Z';
    ];

name = ['DataFile'];

for i=1:size(name,1)
   
dataName=strtok(name(i,:));
bv=strcat(dataName,'.xlsx');
f=xlsread(bv);
[val,ind]=max(f);
[m,n1]=size(ind);

% Connect to Excel
Excel = actxserver('excel.application');
% Get Workbook object
WB = Excel.Workbooks.Open(fullfile(pwd,bv),0,false);

for n=1:n1
count=0;
re=n-26;
while(re ~= 0 & re > 0)
    count=count+1;
    re=floor(re/26);
end
if (abs(n-(26)^(count))<=(26))
    count=count+1;
end
 if count ==2
if(mod(n,26))==0
    secdig='z';
    firdig=alphabet(floor(n/26)-1,:);
     
else
    remi=mod(n,26);
    
    secdig=alphabet(remi,:);
    firdig=alphabet(floor(n/26),:);
end
celname=[firdig,secdig];
 else 
     firdig=alphabet((n),:);
     celname=firdig;
 end
     
cel=[celname,num2str(ind(1,n))];

% Set the color of cell "cel" of Sheet 1 to Yellow
WB.Worksheets.Item(1).Range(cel).Interior.ColorIndex = 6;
end
% Save Workbook
WB.Save();
% Close Workbook
WB.Close();
% Quit Excel
Excel.Quit();
end