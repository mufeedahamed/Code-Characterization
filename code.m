%Save the code and output file in same directory and run the code

source_dir = 'C:\Users\radgr\OneDrive\Desktop\dishan\infiles'; 
filename2 = 'C:\Users\radgr\OneDrive\Desktop\dishan\Table1.xls';
savefilename = 'C:\Users\radgr\OneDrive\Desktop\dishan\output.xlsx'; % Should be the same directory as the code

xlFiles = dir([source_dir, '\*.xls']);
N = length(xlFiles) ;

inimsg = ['Excel files found in directory: ',num2str(N)];
disp(inimsg);

promsg = ['Starting to extract data...'];
disp(promsg);

linemsg = ['-------------------------------------'];
disp(linemsg);

for ii = 1:N
    
thisFile = xlFiles(ii).name ; 
pre_filename = thisFile;  
filename = fullfile(source_dir,pre_filename);

[~, sheets] = xlsfinfo(filename);
[~, sheets2] = xlsfinfo(filename2);

%----------Extract data from sheet POINT--------------------

% Enter your sheet name below
SHEET_NAME_1 = 'POINT';

% Look for cell
search_cell_1_1 = '''borehole';
search_cell_1_2 = 'x (m)';

cell_row_1_1 = [];
cell_col_1_1 = [];
cell_row_1_2 = [];
cell_col_1_2 = [];

[~, ~, raw1] = xlsread(filename, SHEET_NAME_1);
[rowNum_1_1, colNum_1_1] = find(strcmpi(search_cell_1_1, raw1));
[rowNum_1_2, colNum_1_2] = find(strcmpi(search_cell_1_2, raw1));

cell_row_1_1 = [cell_row_1_1; rowNum_1_1];
cell_col_1_1 = [cell_col_1_1; colNum_1_1];
cell_row_1_2 = [cell_row_1_2; rowNum_1_2];
cell_col_1_2 = [cell_col_1_2; colNum_1_2];

f1_1 = raw1(cell_row_1_1 + 1, cell_col_1_1);
f1_2 = raw1{cell_row_1_2 + 1, cell_col_1_2};
f1_3 = raw1{cell_row_1_2 + 1, cell_col_1_2+1};
f1_4 = raw1{cell_row_1_2 + 1, cell_col_1_2+2};
f1_5 = raw1{cell_row_1_2 + 1, cell_col_1_2+3};
f1_6 = raw1{cell_row_1_2 + 1, cell_col_1_2+4};

%disp(f1_4);

f1_5_1 = regexprep(f1_5,'\W','');
f1_1_1 = f1_1(cellfun('isclass',f1_1,'char'));

% Connect to Excel
Excel = actxserver('excel.application');
% Get Workbook object
WB = Excel.Workbooks.Open(fullfile(pwd, 'output.xlsx'), 0, false);
% Get Worksheets object
WS = WB.Worksheets;
% Add after the last sheet
%WS.Add([],WS.Item(WS.Count));
invoke(WS.Item(1),'Copy',[],WS.Item(WS.Count))

namee = char(f1_1_1);
WS.Item(WS.Count).Name = namee;
% Save
WB.Save();
% Quit Excel
Excel.Quit();

format long;
sheetcolhead = {'borehole elevation','borehole depth[m]','GWT depth [m]', 'number of layers','bottom of layers [m]','layers colors','layers types','how to calculate the layers','elevation','depth','SPT','VT remolded [kPa]','VT peak [kPa]','Plasticity Index [%]','Pressumeter test [Mpa]','OCR','UCR [kPa]'};
xlswrite('output.xlsx',sheetcolhead,namee,'A1');
xlswrite('output.xlsx',f1_5_1,namee,'A2');
xlswrite('output.xlsx',f1_2,namee,'A3');
xlswrite('output.xlsx',f1_3,namee,'A4');
xlswrite('output.xlsx',f1_4,namee,'B2');
xlswrite('output.xlsx',f1_6,namee,'C2');

%----------Extract data from sheet LITHOLOGY--------------------

% Enter your sheet name below
SHEET_NAME_2 = 'LITHOLOGY';

% Look for cell
search_cell_2_1 = 'bottom';
search_cell_2_2 = 'uscs';

cell_row_2_1 = [];
cell_col_2_1= [];
cell_row_2_2= [];
cell_col_2_2= [];

[~, ~, raw2] = xlsread(filename, SHEET_NAME_2);
[rowNum_2_1, colNum_2_1] = find(strcmpi(search_cell_2_1, raw2));
[rowNum_2_2, colNum_2_2] = find(strcmpi(search_cell_2_2, raw2));

cell_row_2_1 = [cell_row_2_1; rowNum_2_1];
cell_col_2_1 = [cell_col_2_1; colNum_2_1];
cell_row_2_2 = [cell_row_2_2; rowNum_2_2];
cell_col_2_2 = [cell_col_2_2; colNum_2_2];

f2_1 = raw2(3:end, cell_col_2_1);
f2_2 = raw2(3:end, cell_col_2_2);

f2_2_1 = f2_2(cellfun('isclass',f2_2,'char'));

uniqval = unique(f2_2_1, 'stable');
%disp(uniqval);

[~, ~, raw_tab] = xlsread(filename2, 'Sheet1');
cell_row_tab = [];
cell_col_tab = [];

i = 1;

while i <= length(uniqval)
    [row,column] = find(strcmp(raw2,uniqval(i)));
    b = row';
    b(:);
    D = diff([0,diff(b)==1,0]);
    last = b(D<0);

    for p = 1: length(last)
        colno = cell_col_2_2 - 2;
        lv = raw2(last(p), colno);
        %disp (lv);
        
    end

    if (isempty(last))
        colno = cell_col_2_2 - 2;
        lv = raw2(b, colno);
        %disp (lv);
    end
    
    ccll_1 = sprintf('G%d', i+1);
    ccll_2 = sprintf('E%d', i+1);
    ccll_3 = sprintf('F%d', i+1);
    ccll_4 = sprintf('H%d', i+1);
    xlswrite('output.xlsx',uniqval(i),namee,ccll_1);
    xlswrite('output.xlsx',lv,namee,ccll_2);
    
    [rowNum_tab, colNum_tab] = find(strcmpi(uniqval(i), raw_tab));
    cell_row_tab = [cell_row_tab; rowNum_tab];
    cell_col_tab = [cell_col_tab; colNum_tab];
    
    if isempty(cell_row_tab)
        warmsg = ['Layer value missing for: ',uniqval(i)];
        disp(warmsg);
    else
        %disp(cell_row_tab);
        f_tab_1 = raw_tab(cell_row_tab, cell_col_tab+1);
        f_tab_2 = raw_tab(cell_row_tab, cell_col_tab+2);
        xlswrite('output.xlsx',f_tab_1,namee,ccll_3);
        xlswrite('output.xlsx',f_tab_2,namee,ccll_4);
    end
    cell_row_tab = [];
    cell_col_tab = [];
    
    i = i + 1;
end

%----------Extract data from sheet SPT--------------------

% Enter your sheet name below
SHEET_NAME_3 = 'SPT';

% Look for cell
search_cell_3_1 = 'depth';
search_cell_3_2 = 'nspt';

cell_row_3_1 = [];
cell_col_3_1= [];
cell_row_3_2 = [];
cell_col_3_2= [];

[~, ~, raw3] = xlsread(filename, SHEET_NAME_3);

[rowNum_3_1, colNum_3_1] = find(strcmpi(search_cell_3_1, raw3));
[rowNum_3_2, colNum_3_2] = find(strcmpi(search_cell_3_2, raw3));

cell_row_3_1 = [cell_row_3_1; rowNum_3_1];
cell_col_3_1 = [cell_col_3_1; colNum_3_1];
cell_row_3_2 = [cell_row_3_2; rowNum_3_2];
cell_col_3_2 = [cell_col_3_2; colNum_3_2];

f3_1 = raw3(3:end,cell_col_3_1);
f3_2 = raw3(3:end,cell_col_3_2);

xlswrite('output.xlsx',f3_1,namee,'J2');
xlswrite('output.xlsx',f3_2,namee,'K2');


%----------Extract data from sheet VANE SHEAR--------------------

% Enter your sheet name below
SHEET_NAME_4 = 'VANE SHEAR';

% Look for cell
search_cell_4_1 = 'vane undisturbed';
search_cell_4_2 = 'vane remolded';
search_cell_4_3 = 'depth';

cell_row_4_1 = [];
cell_col_4_1 = [];
cell_row_4_2 = [];
cell_col_4_2 = [];
cell_row_4_3 = [];
cell_col_4_3 = [];

[~, ~, raw4] = xlsread(filename, SHEET_NAME_4);
[rowNum_4_1, colNum_4_1] = find(strcmpi(search_cell_4_1, raw4));
[rowNum_4_2, colNum_4_2] = find(strcmpi(search_cell_4_2, raw4));
[rowNum_4_3, colNum_4_3] = find(strcmpi(search_cell_4_3, raw4));

cell_row_4_1 = [cell_row_4_1; rowNum_4_1];
cell_col_4_1 = [cell_col_4_1; colNum_4_1];
cell_row_4_2 = [cell_row_4_2; rowNum_4_2];
cell_col_4_2 = [cell_col_4_2; colNum_4_2];
cell_row_4_3 = [cell_row_4_3; rowNum_4_3];
cell_col_4_3 = [cell_col_4_3; colNum_4_3];

f4_1 = raw4(3:end,cell_col_4_1);
f4_2 = raw4(3:end,cell_col_4_2);
f4_3 = raw4(3:end,cell_col_4_3);

%disp(f4_1);
%disp(f4_2);
%disp(f4_3);

ff4_1 = cell2mat(f4_1);
ff4_2 = cell2mat(f4_2);

ff4_3 = f4_3(cellfun('isclass',f4_3,'char'));
ff4_3_1 = cell2mat(ff4_3);

usrmsg = ['Data extracted from files: ',num2str(ii),' of ',num2str(N)];
disp(usrmsg);

end

finmsg = ['Data extraction complete.'];
disp(finmsg);
