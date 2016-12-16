clear;
[filename, pathname, filter] = uigetfile('*.xls','Select Excel File');
if filter == 0
    return
end

str = fullfile(pathname,filename);
filename = strcat(pathname,filename);
blsheet = 1;
slsheet = 2;
[bnum,btxt,braw] = xlsread(filename, blsheet);
[snum,stxt,sraw] = xlsread(filename, slsheet);

result = sraw;
[sheight, swidth] = size(sraw);
[bheight, bwidth] = size(braw);
prog = sheight / 100;
for i=2:sheight
    swellnum = sraw{i, 1};
    stop = sraw{i, 3};
    sbtn = sraw{i, 4};
    pos = -1;
    name = '';
    for j=2:bheight
        bwellnum = braw{j, 1};
        btop = braw{j, 3};
        bbtn = braw{j, 4};
        if strcmp(bwellnum,swellnum) && (stop >= btop) && (sbtn <= bbtn)
            pos = j;
            name = braw{j, 2};
            break;
        end
    end
    if pos ~= -1
        result{i, 2} = name;
    end
    
    if i >= prog;
        prog = prog + sheight / 100;
        fprintf('%.0f/100\n', i/sheight * 100);
    end
end

[filename, pathname, filter] = uiputfile('*.xls', 'Save result');
if filter == 0
    return
end

filename = strcat(pathname,filename);

xlswrite(filename, result);