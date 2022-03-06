clear;
clc;

disp('请选择Excel文件...')
%[filename, pathname] = uigetfile({'*.xls';'*.xlsx'},' Select the Data Dictionary');
[filename, pathname] = uigetfile({'*.xlsx'},' Select the Data Dictionary');

if isequal(filename,0)
    disp('取消操作')
    return
else
    disp(['数据文件路径：',fullfile(pathname,filename)])
end

disp('正在处理...')

[~,name,~] = fileparts(filename);
file_name = name;

Barindex = 0;
h = waitbar(0,'正在读取数据......');

fid = fopen('DataObject.m','w');

[~,DataType_Text] = xlsread(file_name,'DataType');
[datatype_num,~] = size(DataType_Text);

[~,StructBus_Text]=xlsread(file_name,'Struct');
[structbus_num,~] = size(StructBus_Text);

[~,VarCal_Text] = xlsread(file_name,'Parameter');
[varcal_num,~] = size(VarCal_Text);

[~,VarGlobal_Text] = xlsread(file_name,'Signal');
[varglobal_num,~] = size(VarGlobal_Text);

BarNum = datatype_num + structbus_num + varcal_num +varglobal_num;

fprintf(fid,'clear;');
fprintf(fid,'Barindex =0;');
fprintf(fid,'\n');
str_tmp =strcat('BarNum =',num2str(BarNum),';');
fprintf(fid,'\n');
fprintf(fid,'h=waitbar(0,''正在定义数据......'');');
fprintf(fid,str_tmp);
fprintf(fid,'\n\n');
for row = 2:datatype_num
    Object_TypeName = char(DataType_Text(row,1));
    Object_Type = char(DataType_Text(row,2));
    Object_Scope = char(DataType_Text(row,3));
    Object_Header = char(DataType_Text(row,4));

    if ~isempty(Object_TypeName)
    str_tmp = strcat(Object_TypeName,'=Simulink.AliasType;');
    fprintf(fid,str_tmp);
    fprintf(fid,'\n');

    str_tmp = strcat(Object_TypeName,'.DataScope =''',Object_Scope,''',');
    fprintf(fid,str_tmp);
    fprintf(fid,'\n');

    str_tmp = strcat(Object_TypeName,'.HeaderFile =''',Object_Header,''',');
    fprintf(fid,str_tmp);
    fprintf(fid,'\n');
    
    str_tmp = strcat(Object_TypeName,'.BaseType =''',Object_Type,''',');
    fprintf(fid,str_tmp);
    fprintf(fid,'\n\n');
    
    Barindex = Barindex + 1;
    waitbar(Barindex/BarNum);
    end
end

saveVarsNum = 1;
for row = 2:structbus_num
    Barindex = Barindex + 1;
    waitbar(Barindex/BarNum);
    Object_VarName = char(StructBus_Text(row,1));
    Object_Scope = char(StructBus_Text(row,2));
    Object_Elem = char(StructBus_Text(row,3));
    Object_Bus = char(StructBus_Text(row,4));
    Object_Header = char(StructBus_Text(row,5));
    Object_Type = char(StructBus_Text(row,6));
    Object_Dim = char(StructBus_Text(row,9));
    if ~isempty(Object_VarName)
        if strcmp(Object_Elem,'Bus')
            str_tmp = strcat(Object_VarName,' = Simulink.Bus;');
            fprintf(fid,str_tmp);
            fprintf(fid,'\n');
            
            str_tmp =strcat(Object_VarName,'.DataScope=''',Object_Scope,''';');
            fprintf(fid,str_tmp);
            fprintf(fid,'\n');
            
            strtmp =strcat(Object_VarName,'.HeaderFile =''',Object_Header,''';');
            fprintf(fid,str_tmp);
            fprintf(fid,'\n');
        else
            str_tmp = strcat('saveVarsTmp{1}(',num2str(saveVarsNum),',1) = Simulink.BusElement;');
            fprintf(fid,str_tmp);
            fprintf(fid,'\n');
            
            str_tmp=strcat('saveVarsTmp{1}(',num2str(saveVarsNum),',1).Name = ''',Object_VarName,''';');
            fprintf(fid,str_tmp);
            fprintf(fid,'\n');
            
            str_tmp = strcat('saveVarsTmp{1}(',num2str(saveVarsNum),',1).Complexity = ''real'';');
            fprintf(fid,str_tmp);
            fprintf(fid,'\n');
            
            str_tmp =strcat('saveVarsTmp{1}(',num2str(saveVarsNum),',1).Dimensions = ',Object_Dim,';');
            fprintf(fid,str_tmp);
            fprintf(fid,'\n');
            
            str_tmp=strcat('saveVarsTmp{1}(',num2str(saveVarsNum),',1).DataType = ''',Object_Type,''';');
            fprintf(fid,str_tmp);
            fprintf(fid,'\n');
            
            saveVarsNum = saveVarsNum + 1;
            
            if row < structbus_num
                row_next = row + 1;
                Object_ElemNext = char(StructBus_Text(row_next,3));
                
                if strcmp(Object_ElemNext,'Bus')
                    saveVarsNum = 1;
                    str_tmp = strcat(Object_Bus,'.Elements = saveVarsTmp{1};');
                    fprintf(fid,str_tmp);
                    fprintf(fid,'\n');
                    
                    fprintf(fid,'clear saveVarsTmp;');
                    fprintf(fid,'\n\n');
                    
                    fprintf(fid,'Barindex = Barindex + 1;');
                    fprintf(fid,'\n');
                    fprintf(fid,'waitbar(Barindex/BarNum);');
                    fprintf(fid,'\n\n');
                end
            else
                str_tmp =strcat(Object_Bus,'.Elements = saveVarsTmp{1};');
                fprintf(fid,str_tmp);
                fprintf(fid,'\n');
                
               fprintf(fid,'clear saveVarsTmp;');
               fprintf(fid,'\n\n');
               
               fprintf(fid,'Barindex = Barindex + 1;');
               fprintf(fid,'\n');
               fprintf(fid,'waitbar(Barindex/BarNum);');
               fprintf(fid,'\n\n');
            end
            
        end
    end
end

for row = 2:varglobal_num
    Barindex = Barindex + 1; 
    waitbar(Barindex/BarNum);
    Object_VarName = char(VarGlobal_Text(row,1)); 
    Object_Package = char(VarGlobal_Text(row,2));
    Object_Object = char(VarGlobal_Text(row,3)); 
    Object_Storage = char(VarGlobal_Text(row,4));
    Object_Type = char(VarGlobal_Text(row,5));
    Object_Dim = char(VarGlobal_Text(row,9));
    Object_Initial = char(VarGlobal_Text(row,12));
    Object_Description = char(VarGlobal_Text(row,13));
    
    if ~isempty(Object_VarName)
        str_tmp = strcat(Object_VarName,'=',32,Object_Package,'.',Object_Object,';');  %Var_Name = Simulink.Signal;
        fprintf(fid,str_tmp);
        fprintf(fid,'\n');
    
        str_tmp = strcat(Object_VarName,'.CoderInfo.StorageClass =''',Object_Storage,''';');
        fprintf(fid,str_tmp);
        fprintf(fid,'\n');
    
        str_tmp = strcat(Object_VarName,'.Complexity =''real'';');
        fprintf(fid,str_tmp);
        fprintf(fid,'\n');
    
        str_tmp = strcat(Object_VarName,'.InitialValue =''',Object_Initial,''';');
        fprintf(fid,str_tmp); 
        fprintf(fid,'\n');
    
        str_tmp = strcat(Object_VarName,'.Dimensions =',32,Object_Dim,';');
        fprintf(fid,str_tmp);
        fprintf(fid,'\n');
    
        str_tmp = strcat(Object_VarName,'.DataType =''',Object_Type,''';');
        fprintf(fid,str_tmp);
        fprintf(fid,'\n');
    
        str_tmp = strcat(Object_VarName,'.Description =''',Object_Description,''';');
        fprintf(fid,str_tmp);
        fprintf(fid,'\n\n');
    
        fprintf(fid,'Barindex = Barindex + 1;');
        fprintf(fid,'\n');
        fprintf(fid,'waitbar(Barindex/BarNum);'); 
        fprintf(fid,'\n\n');
    end
    
end

for row = 2:varcal_num
    Barindex = Barindex + 1;
    waitbar(Barindex/BarNum);
    Object_VarName = char(VarCal_Text(row,1));
    Object_Package = char(VarCal_Text (row,2));
    Object_Object = char(VarCal_Text(row,3));
    Object_Storage = char(VarCal_Text(row,4));
    Object_Type = char(VarCal_Text (row,5));
    Object_Initial = char(VarCal_Text (row,12));
    if ~isempty(Object_VarName)
        str_tmp = strcat(Object_VarName,'=',32,Object_Package,'.',Object_Object,';');  %Var_Name = Simulink.Parameter
        fprintf(fid,str_tmp);
        fprintf(fid,'\n');
        
        str_tmp = strcat(Object_VarName,'.Value =',32,Object_Initial,';');
        fprintf(fid,str_tmp);
        fprintf(fid,'\n');
        
        str_tmp = strcat(Object_VarName,'.CoderInfo.StorageClass =''',Object_Storage,''';');
        fprintf(fid,str_tmp);
        fprintf(fid,'\n');
        
        str_tmp = strcat(Object_VarName,'.DataType=''',Object_Type,''';');
        fprintf(fid,str_tmp);
        fprintf(fid,'\n\n');
        
        fprintf(fid,'Barindex = Barindex+1;');
        fprintf(fid,'\n');
        fprintf(fid,'waitbar(Barindex/BarNum);');
        fprintf(fid,'\n');
        
        fprintf(fid,'\n\n');
        
    end
end

close(h);
clear h;

fprintf(fid,'close(h);');
fprintf(fid,'\n');
fprintf(fid,'clear h;');
fprintf(fid,'\n');
fprintf(fid,'clear Barindex;');
fprintf(fid,'\n');
fprintf(fid,'clear BarNum;');
fprintf(fid,'\n\n');

disp('正在定义数据...')
fclose(fid);
clear;clc;
%DataObject;
disp('处理完成！')

[DataDictName,~] = uiputfile('*.sldd','Save as a data dictionary file');

if DataDictName ~= 0
    exist_flag = exist(DataDictName,'file');
    
    if exist_flag == 0
        NewDictObj = Simulink.data.dictionary.create(DataDictName);
    else
        NewDictObj = Simulink.data.dictionary.open(DataDictName);
    end
    
    [FileName,~,~] = uigetfile('*.m','Import from Matlab file');
    if FileName ~= 0
        DataSectObj = getSection(NewDictObj,'Design Data');
        ImportedVars = importFromFile(DataSectObj,FileName,'existingVarsAction','overwrite');
        saveChanges(NewDictObj);
        
        msgbox('New sldd file has been created.');
    end
end
    