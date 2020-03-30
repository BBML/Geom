%% Average Plot Creator
% Get data
clear all

[prof_filename, prof_pathname] = uigetfile({'*.xls;*.xlsx;*.csv','Excel Files (*.xls,*.xlsx,*.csv)'; '*.*',  'All Files (*.*)'},'Pick the file with CT_prof info');
prof = xlsread([prof_pathname prof_filename],'Raw Data');
[~,names] = xlsread([prof_pathname prof_filename],'Raw Data');

[key_filename, key_pathname] = uigetfile({'*.xls;*.xlsx;*.csv','Excel Files (*.xls,*.xlsx,*.csv)'; '*.*',  'All Files (*.*)'},'Pick the file with key info');
[~,~,key] = xlsread([key_pathname key_filename],1);


%% Define names for the test categories and their corresponding variables
% Overall categories
cat_1=key{1,2};
cat_2=key{1,3};

% Variables in category 1
cat1_1=key{2,2};
for j=2:length(key)
    if strcmp(cat1_1,key{j,2})
    else
        cat1_2=key{j,2};
        break
    end
end

%Variables in category 2
cat2_1=key{2,3};
for j=2:length(key)
    if strcmp(cat2_1,key{j,3})
    else
        cat2_2=key{j,3};
        break
    end
end

A_name=[cat1_1 ' ' cat2_1];
B_name=[cat1_1 ' ' cat2_2];
C_name=[cat1_2 ' ' cat2_1];
D_name=[cat1_2 ' ' cat2_2];

%% Create the four vector sets
% Set counters for each set
ca=0;
cb=0;
cc=0;
cd=0;

% Create sets of ID names sorted into four groups based on defined variables
for j=2:length(key)
    if strcmp(key{j,2},cat1_1) && strcmp(key{j,3},cat2_1)
        ca=ca+1;
        A{ca}=key{j,1};
    elseif strcmp(key{j,2},cat1_1) && strcmp(key{j,3},cat2_2)
        cb=cb+1;
        B{cb}=key{j,1};
    elseif strcmp(key{j,2},cat1_2) && strcmp(key{j,3},cat2_1)
        cc=cc+1;
        C{cc}=key{j,1};
    else
        cd=cd+1;
        D{cd}=key{j,1};
    end
end

%% Calculate average profiles
% Create master set of the sorted specimen lists
sorted={A_name A;B_name B;C_name C;D_name D};

% Cycle through each group
for i=1:4
    group=sorted{i,2};
    name=sorted{i,1};
    
    % Collect all 
    for j=1:length(group)
        ID=group{j};
        data_place=find(strcmp(names,ID));
        data_1c(j,:)=prof(data_place(1),:);
        data_2c(j,:)=prof(data_place(2),:);
    end
    
    data_1=mean(data_1c);
    data_2=mean(data_2c);
    results{i,:}={name, data_1, data_2};
end
        
% Finally, create plots of average profiles     
% Plot 1
theta=prof(1,:)*pi/180;
A_set=results{1};
B_set=results{2};
[x1,y1]=pol2cart(theta,A_set{2});
[x2,y2] = pol2cart(theta,A_set{3});
[x3,y3]=pol2cart(theta,B_set{2});
[x4,y4] = pol2cart(theta,B_set{3});

figure(1)
patch([x1 fliplr(x2)], [y1 fliplr(y2)], 'b', 'EdgeColor','k','FaceAlpha',.3)
hold on
plot(x3, y3, '-r')
plot(x4, y4, '-r')
set(gca,'ytick',[])
set(gca,'xtick',[])
hold off
legend(A_name,B_name)
print ('-dpng', [cat1_1 'comp']) 

% Plot 2
C_set=results{3};
D_set=results{4};
[x1,y1]=pol2cart(theta,C_set{2});
[x2,y2] = pol2cart(theta,C_set{3});
[x3,y3]=pol2cart(theta,D_set{2});
[x4,y4] = pol2cart(theta,D_set{3});

figure(2)
patch([x1 fliplr(x2)], [y1 fliplr(y2)], 'b', 'EdgeColor','k','FaceAlpha',.3)
hold on
plot(x3, y3, '-r')
plot(x4, y4, '-r')
set(gca,'ytick',[])
set(gca,'xtick',[])
hold off
legend(C_name,D_name)
print ('-dpng', [cat1_2 'comp']) 