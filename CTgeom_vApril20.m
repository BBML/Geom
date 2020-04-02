function [ ]  = CTgeom_fc()
%% Revision History

% Edited April 2020 by Rachel Kohler to change the input method and the
% output files by introducing a key file in place of a folder heiarchy.
% This enables the code to pre-sort the data into the testing groups, and
% output organized data.

% Edited Oct 3 2019 by Rachel Kohler to output c_ant (anterior extreme).
% Further edits done Oct 10 2019 to clean up code (deleting unused code).

% Edited May 2015 by Max Hammond to optimize code, record a diary,
% calculate TMD, read BMD equations from multiple sub-folders, reduce Excel
% write functions, removed xlswrite1 dependence, fill pores, apply
% greyscale threshold, remove naming dependence, alter parameter output
% order, and alter figure output. Outdated code is commented out. Note that
% once code is commented out it may not function as expected if re-inserted
% given subsequent changes.

% Edited Sept 2014 by Max Hammond to add code from Alycia Berman that
% excludes scale bars in the images at line 105. Changed outpout to xlsx
% format and added headers. Adjust profile output to go from 0 to 360.
% Added studynum so the loop doesn't have to be hard coded. Used xlswrite1
% which only opens Excel once per file to speed up program. This adds a
% dependancy to xlswrite1.m which must be in the directory to run the file

% Written by Joey Wallace, Sept 2011

%% Function Overview

% This program reads grayscale .bmp images from CT, applies a theshold,
% removes everything except the bone, fills in pores, and calculates
% relevant geometric properties. Each slice is calculated individually.
% Four Excel spreadsheets are output containing slice by slice or average
% geometric properties or profiles. A .png image is output for each bone
% within the respective folder showing the bone's profile and major/minor
% axes.

% If there an error is generated during analysis, the program will save the
% data already analyzed if any and display a warning with information on
% when the error occured along with the actual error message.

%% Proper Setup
% These files should be vertically aligned and oriented with anterior to the
% right.  Therefore, right limbs will be oriented medial up and left limbs
% will be oriented medial down.

% FOLDER HEIRARCHY SETUP IS NO LONGER NEEDED
% All folders of .bmp files go in the same folder as this MATLAB code. To
% create the required key file, make an Excel file with the following info:
% 
% KEY FILE
% Column 1- specimen IDs (must match the folder names)
% Column 2- test variable type 1 (E.g., Sex: Male or Female)
% Column 3- test variable type 2 (E.g., Genotype: WT or Amish)
% Column 4- AC-# value (Phantom info from that specimen's scan day)
% Column 5- Denominator (Phantom info from that specimen's scan day)

% There is no longer a naming convention that you must follow. Each bone
% can have a different number of slices, but all bones must have the sample
% resolution, be the same bone, and be from the same side.

%% Initialization
diary off
close all % close all figures
clear all % clear all variables
format long % change format to long
warning('off','MATLAB:xlswrite:AddSheet'); % disable add sheet warning for Excel

%% Import a key file with the test variable info for each specimen ID, sorted by columns
[key_filename, key_pathname] = uigetfile({'*.xls;*.xlsx;*.csv','Excel Files (*.xls,*.xlsx,*.csv)'; '*.*',  'All Files (*.*)'},'Pick the file with key info');
[~,~,key] = xlsread([key_pathname key_filename],1);

% Define names for the test categories and their corresponding variables
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

% Create the four vector sets
% Set counters for each set
ca=0;
cb=0;
cc=0;
cD=0;

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
        cD=cD+1;
        D{cD}=key{j,1};
    end
end

sorted={A_name A;B_name B;C_name C;D_name D};

%% Input parameters and check for common erros

res = input ('\n\n Voxel resolution in um: '); % input the isotropic voxel size from CTAn

% Input step used for extreme calculations (typically 0.5) and display a
% warning for an odd angular resolution
ang = input (' Angle step in degrees (factor of 360): ');
while mod(360,ang)
    fprintf(2,'\n Please enter a factor of 360 (e.g. 0.5). Note: this is NOT the angular step size used in the uCT scan. \n\n')
    ang = input ('Angle step in degrees (factor of 360): ');
end

% Input global greyscale threshold and display a warning for values outside
% of the specified range
threshold = input(' Input threshold value (0-255): ');
while threshold>255 || threshold<0
    fprintf(2,'\n Please enter a threshold between 0 and 255. \n\n')
    threshold = input('Input threshold value (0-255): ');
end

% Input whether left or right limbs will be analyzed, correct common
% mistakes, and display a warning for other values
side = input(' Enter "l" for a left limb and "r" for a right limb: ', 's');
side = strrep(side,'"','');
side = lower(side);
side = side(1:1);
side_log = ~strcmp(side,'l') + ~strcmp(side,'r');
while side_log~=1
    fprintf(2, ['\n You entered "' side '". \n\n'])
    side = input(' Enter "l" for a left limb and "r" for a right limb: ', 's');
    side = strrep(side,'"','');
    side = lower(side);
    side = side(1:1);
    side_log = ~strcmp(side,'l') + ~strcmp(side,'r');
end

% Input whether femora or tibiae will be analyzed, correct common mistakes,
% and display a warning for other values
bone = input(' Enter "f" for a femur and "t" for a tibia: ', 's');
bone = strrep(bone,'"','');
bone = lower(bone);
bone = bone(1:1);
bone_log = ~strcmp(bone,'f') + ~strcmp(bone,'t');
while bone_log~=1
    fprintf(2, ['\n You entered "' bone '". \n\n']) 
    bone = input(' Enter "f" for a femur and "t" for a tibia: ', 's');
    bone = strrep(bone,'"','');
    bone = lower(bone);
    bone = bone(1:1);
    bone_log = ~strcmp(bone,'f') + ~strcmp(bone,'t');
end
% tot = tic; % start the stop watch

%% Calculation
% Preallocate variables
A = 360/ang+1;
ac_step = .11/255;
offset = 0;
ppp=1;

% Cycle through each group
for k=1:4
    group=sorted{k,2};
    name=sorted{k,1};
    count=0;
    prof_peri_avg=zeros(1,721);
    prof_endo_avg=zeros(1,721);
    % Cycle through specimens in each group
    for m=1:length(group)
        filename=group{m};
        if exist(filename, 'dir')~=7
            fprintf('Folder not found for %s.\n',filename)
        else
            fprintf('Processing %s.\n',filename)
            count=count+1;
%             tic
%     clearvars -except key group group_list name filename ppp offset res tot ang side bone threshold A prof_out_peri_cell prof_out_endo_cell sample_list geom_out_cell j i ac_step
    
    % Get info for this specimen
        sample_list{ppp} = filename;
        group_list{ppp}=name;
        data_place=find(strcmp(key,filename));
        eq_num=key{data_place,4};
        eq_denom=key{data_place,5};
        
    % Store the .bmp filenames in the folder
        cd(filename);
        slices=dir('*.bmp');
        
        % Pre-allocating arrays for data output later
        peri_out = zeros(length(slices),A);
        endo_out = zeros(length(slices),A);
        profiles = zeros((2*length(slices)+3),A);
        geom_out = zeros(length(slices)+2,19);
        tmd_gs = zeros(length(slices), 3);
        centroid_x = zeros(length(slices), 1);
        centroid_y = zeros(length(slices), 1);
            
            % Create loop to calculate parameters for each slice
            for j=1:length(slices)
                
                % Read in each slice as a grayscale image
                section = imread(slices(j,1).name);
                
                % Read in each slice as a BW image and allow the variable
                % to change size to accommodate different sized ROIs and/or
                % different number of slices
  
                slice = imbinarize(section,(threshold-1)/255); % use threshold-1 to take everything greater than or equal to the input threshold like CTAn
                
                % Remove all but the largest connected component (i.e.
                % remove scales or fibula if applicable)
                cc=bwconncomp(slice); % find the connected components in the image
                numPixels = cellfun(@numel,cc.PixelIdxList); % find the number of pixels in each component
                [~,idx] = max(numPixels); % find the index containing the most number of pixels
                for i=1:length(cc.PixelIdxList) % remove all other components
                    if i==idx
                        % do nothing
                    else
                        slice(cc.PixelIdxList{i})=0;
                    end
                end
                clear cc numPixels idx i
                
                stats = regionprops(slice,section,'MeanIntensity','PixelValues');
                tmd_gs(j,:) = [stats.MeanIntensity sum([stats.PixelValues]) length([stats.PixelValues])];
                tmd_ac = tmd_gs(1).*ac_step;
                tmd_HA = (tmd_ac - eq_num)/eq_denom;
                
                % Manually calculate the centroid from pixel locations:
                [index_y,index_x] = find (slice == 1); % this finds the x and y locations of each "on" pixel
                Qx = sum(index_y); % since the area of each dA is 1 pixel, we have x1 + x2 +...+ and this is the integral of y_dA
                Qy = sum(index_x); % since the area of each dA is 1 pixel, this is the integral of x_dA
                area = length(index_y); % since the area of each pixel is one, this is the total number of pixels or the area
                xbar = Qy/area; % xbar = integral of y_da/A
                ybar = Qx/area; % ybar = integral of x_da/A
                
                % PLOT THE IMAGE WITH THE CENTROID MARKED
                % Start getting line profiles at various degrees:
                inner_fiber = zeros(1, A); % creats a zero vector for the endocortical radii
                outer_fiber = zeros(1, A); % creates a zero vector for the periosteal radii
                thickness = zeros(1, A-1);  % creates a zero vector for cortical thicknesses
                index = 0; % intialize index
                
                image_size = size(slice);
                x_size = image_size(2);
                y_size = image_size(1);
                
                % IN QUADRANT 1 (45 to 134.99 degrees, top quadrant):
                for i = -45:ang:44.9
                    index = index + 1;
                    angle = i * pi / 180;
                    yi=[ybar,0];
                    xi=[xbar,xbar-(tan(angle)*yi(1))];
                    [cx,cy,c] = improfile(slice,xi,yi,10000);
                    % improfile chooses an arbirtray number of points to
                    % look at so I will choose alot to be accurate. cx and
                    % cy are the pixel locations along the line and c is
                    % the intensity at each point
                    cort_on = find(c == 1);
                    thick = length(cort_on);
                    on_1 = cort_on(1);
                    on_end = cort_on(thick);
                    rad_x = [cx(on_1),cx(on_end)];
                    rad_y = [cy(on_1),cy(on_end)];
                    points=[xbar,ybar;rad_x(1),rad_y(1);rad_x(2),rad_y(2)]; % centroid, endo and peri points along this line
                    radii = pdist(points); % radii(1) is endo, radii(2) is peri and radii(3) is c_thickness
                    inner_fiber(1, index) = radii(1);
                    outer_fiber(1, index) = radii(2);
                    thickness(1, index) = radii(3);
                end
                
                
                % IN QUADRANT 2 (135 to 224.99 degrees, P quadrant):
                for i = -45:ang:44.9
                    index = index + 1;
                    angle = i * pi / 180;
                    xi=[xbar,0];
                    yi=[ybar,ybar+(tan(angle)*xi(1))];
                    [cx,cy,c] = improfile(slice,xi,yi,10000);
                    cort_on = find(c == 1);
                    thick = length(cort_on);
                    on_1 = cort_on(1);
                    on_end = cort_on(thick);
                    rad_x = [cx(on_1),cx(on_end)];
                    rad_y = [cy(on_1),cy(on_end)];
                    points=[xbar,ybar;rad_x(1),rad_y(1);rad_x(2),rad_y(2)];
                    radii = pdist(points);
                    inner_fiber(1, index) = radii(1);
                    outer_fiber(1, index) = radii(2);
                    thickness(1, index) = radii(3);
                end
                
                % IN QUADRANT 3 (225 to 314.99 degrees, bottom quadrant):
                for i = -45:ang:44.9
                    index = index + 1;
                    angle = i * pi / 180;
                    yi=[ybar,y_size];
                    xi=[xbar,xbar+(tan(angle)*(yi(2)-yi(1)))];
                    [cx,cy,c] = improfile(slice,xi,yi,10000);
                    cort_on = find(c == 1);
                    thick = length(cort_on);
                    on_1 = cort_on(1);
                    on_end = cort_on(thick);
                    rad_x = [cx(on_1),cx(on_end)];
                    rad_y = [cy(on_1),cy(on_end)];
                    points=[xbar,ybar;rad_x(1),rad_y(1);rad_x(2),rad_y(2)];
                    radii = pdist(points);
                    inner_fiber(1, index) = radii(1);
                    outer_fiber(1, index) = radii(2);
                    thickness(1, index) = radii(3);
                end
                
                % IN QUADRANT 4 (315 to 404.99 or 49.99, A quadrant):
                for i = -45:ang:44.9
                    index = index + 1;
                    angle = i * pi / 180;
                    xi=[xbar,x_size];
                    yi=[ybar,ybar-(tan(angle)*(xi(2)-xi(1)))];
                    [cx,cy,c] = improfile(slice,xi,yi,10000);
                    cort_on = find(c == 1);
                    thick = length(cort_on);
                    on_1 = cort_on(1);
                    on_end = cort_on(thick);
                    rad_x = [cx(on_1),cx(on_end)];
                    rad_y = [cy(on_1),cy(on_end)];
                    points=[xbar,ybar;rad_x(1),rad_y(1);rad_x(2),rad_y(2)];
                    radii = pdist(points);
                    inner_fiber(1, index) = radii(1);
                    outer_fiber(1, index) = radii(2);
                    thickness(1, index) = radii(3);
                end
                
                % To plot this in polar,you need to append the inner and
                % outer vectors with the value at 360 degrees (0 deg) to
                % close
                index = index + 1;
                inner_fiber(1, index) = inner_fiber(1);
                outer_fiber(1, index) = outer_fiber(1);
                
                % Setup angles for polar plot
                angle_deg = 45:ang:405;
                angle_rad = angle_deg.*pi./180; %convert to radians

                % Convert the geometric data from angle and radius to x and
                % y coordinates
                outer_fiber_x = outer_fiber.*cos(angle_rad);
                outer_fiber_y = outer_fiber.*sin(angle_rad);
                inner_fiber_x = inner_fiber.*cos(angle_rad);
                inner_fiber_y = inner_fiber.*sin(angle_rad);
                
                % Before shifting the origin from the centroid, calculate
                % extreme fiber in each anatomic direction.  A and P are
                % not dependent on wheter this is a right or left bone, but
                % M and L are:
                anterior_extreme = abs(max(outer_fiber_x));
                posterior_extreme = abs(min(outer_fiber_x));
                
                if side == 'r'
                    medial_extreme = abs(max(outer_fiber_y));
                    lateral_extreme = abs(min(outer_fiber_y));
                else
                    medial_extreme = abs(min(outer_fiber_y));
                    lateral_extreme = abs(max(outer_fiber_y));
                end
                
                % Shift the coordinate system from (0,0) at centroid to the
                % (0,0) at LL corner.  For geometric properties, the outer
                % perimeter needs to go in the CW direction.  Currently, it
                % is CCW so it needs to be flipped.  This does both and
                % plots to verify
                x_data_min = abs(min(outer_fiber_x));
                y_data_min = abs(min(outer_fiber_y));
                outer_fiber_x = outer_fiber_x+x_data_min;
                outer_fiber_x = fliplr(outer_fiber_x);
                outer_fiber_y = outer_fiber_y+y_data_min;
                outer_fiber_y = fliplr(outer_fiber_y);
                inner_fiber_x = inner_fiber_x+x_data_min;
                inner_fiber_y = inner_fiber_y+y_data_min;
                x_perimeter = [outer_fiber_x inner_fiber_x];
                y_perimeter = [outer_fiber_y inner_fiber_y];
                
                % USE THESE OUTPUTS PRIOR TO INCORPORATING POLYGEOM TO GET
                % SOME OF THE GEOMETRIC PROPERTIES OF INTEREST
                
                % Convert all pixel values to um and plot
                outer_fiber_x = outer_fiber_x*res;
                outer_fiber_y = outer_fiber_y*res;
                inner_fiber_x = inner_fiber_x*res;
                inner_fiber_y = inner_fiber_y*res;
                x_perimeter = x_perimeter*res;
                y_perimeter = y_perimeter*res;
                
                subplot(1,2,2)
                plot(x_perimeter,y_perimeter)
                hold on
                axis equal
                axis tight
                xlabel('x-position in um')
                ylabel('y-position in um')
                title('Calculated Profiles')
                
                % Calculate the average cortical thickness
                cort_thickness = mean(thickness);
                cort_thickness = cort_thickness*res;
                
                % Calculate extreme fiber in each anatomic direction in um
                anterior_extreme = anterior_extreme * res;
                medial_extreme = medial_extreme * res;
                posterior_extreme = posterior_extreme * res;
                lateral_extreme = lateral_extreme * res;
                
                % Calculate AP and ML diameters and AP/ML
                AP_width = anterior_extreme + posterior_extreme;
                ML_width = medial_extreme + lateral_extreme;
                APtoMLratio = AP_width/ML_width;
                
                %*****ADD POLYGEOM TO GET TOTAL CROSS SECTIONAL AREA AND
                %PERISOTEAL PERIMETER*****
                clear x y
                x = outer_fiber_x;
                y = outer_fiber_y;
                
                % Check if inputs are same size
                if ~isequal( size(x), size(y) )
                    error( 'X and Y must be the same size');
                end
                
                % Number of vertices
                [ x, ~ ] = shiftdim( x ); 
                [ y, ~ ] = shiftdim( y ); 
                [ n, ~ ] = size( x ); 
                
                % Temporarily shift data to mean of vertices for improved
                % accuracy
                xm = mean(x);
                ym = mean(y);
                x = x - xm*ones(n,1);
                y = y - ym*ones(n,1);
                
                % Delta x and delta y
                dx = x ( [ 2:n 1 ] ) - x;
                dy = y ( [ 2:n 1 ] ) - y;
                
                % Output of Total Cross Sectional Area
                total_cs_area = sum( y.*dx - x.*dy )/2;
                periosteal_BS = sum( sqrt( dx.*dx +dy.*dy ) );
                
                
                %*****ADD PART OF POLYGEOM TO GET CORTICAL AND MARROW
                %AREAS*****
                clear x y
                x = fliplr(inner_fiber_x);
                y = fliplr(inner_fiber_y);
                
                % check if inputs are same size
                if ~isequal( size(x), size(y) )
                    error( 'X and Y must be the same size');
                end
                
                % Number of vertices
                [ x, ~ ] = shiftdim( x ); 
                [ y, ~ ] = shiftdim( y ); 
                [ n, ~ ] = size( x ); 
                
                % Temporarily shift data to mean of vertices for improved
                % accuracy
                xm = mean(x);
                ym = mean(y);
                x = x - xm*ones(n,1);
                y = y - ym*ones(n,1);
                
                % Delta x and delta y
                dx = x ( [ 2:n 1 ] ) - x;
                dy = y ( [ 2:n 1 ] ) - y;
                
                % Output of Marrow Area and Cortical Area
                marrow_area = sum( y.*dx - x.*dy )/2;
                cortical_area = total_cs_area - marrow_area;
                endocortical_BS = sum( sqrt( dx.*dx +dy.*dy ) );
                
                %*****NOW INCORPORATE POLYGEOM TO GET OTHER PROPS*****
                clear x y
                x = x_perimeter;
                y = y_perimeter;
                
                % Check if inputs are same size
                if ~isequal( size(x), size(y) )
                    error( 'X and Y must be the same size');
                end
                
                % Number of vertices
                [ x, ~ ] = shiftdim( x ); 
                [ y, ~ ] = shiftdim( y ); 
                [ n, ~ ] = size( x ); 
                
                % Temporarily shift data to mean of vertices for improved
                % accuracy
                xm = mean(x);
                ym = mean(y);
                x = x - xm*ones(n,1);
                y = y - ym*ones(n,1);
                
                % Delta x and delta y
                dx = x( [ 2:n 1 ] ) - x;
                dy = y( [ 2:n 1 ] ) - y;
                
                % Summations for CW boundary integrals
                cA = sum( y.*dx - x.*dy )/2; % cortical area
                Axc = sum( 6*x.*y.*dx -3*x.*x.*dy +3*y.*dx.*dx +dx.*dx.*dy )/12; % first moment about the y-axis (xc*cA)
                Ayc = sum( 3*y.*y.*dx -6*x.*y.*dy -3*x.*dy.*dy -dx.*dy.*dy )/12; % first moment about the x-axis (yc*cA)
                Ixx = sum( 2*y.*y.*y.*dx -6*x.*y.*y.*dy -6*x.*y.*dy.*dy ...
                    -2*x.*dy.*dy.*dy -2*y.*dx.*dy.*dy -dx.*dy.*dy.*dy )/12;% second moment about x axis
                Iyy = sum( 6*x.*x.*y.*dx -2*x.*x.*x.*dy +6*x.*y.*dx.*dx ...
                    +2*y.*dx.*dx.*dx +2*x.*dx.*dx.*dy +dx.*dx.*dx.*dy )/12;% second moment about y axis
                Ixy = sum( 6*x.*y.*y.*dx -6*x.*x.*y.*dy +3*y.*y.*dx.*dx ...
                    -3*x.*x.*dy.*dy +2*y.*dx.*dx.*dy -2*x.*dx.*dy.*dy )/24;% product of inertia about x-y axes
                %P = sum( sqrt( dx.*dx +dy.*dy ) ); % perimeter?
                
                % Check for CCW versus CW boundary
                if cA < 0
                    cA = -cA;
                    Axc = -Axc;
                    Ayc = -Ayc;
                    Ixx = -Ixx;
                    Iyy = -Iyy;
                    Ixy = -Ixy;
                end
                
                % Centroidal moments
                xc = Axc / cA; % centroidal location in x direction
                yc = Ayc / cA; % centroidal location in y direction
                Iuu = Ixx - cA*yc*yc; % centroidal MOI about x axis
                Ivv = Iyy - cA*xc*xc; % centroidal MOI anout y axis
                Iuv = Ixy - cA*xc*yc; % product of inertia

                % Replace mean of vertices
                x_cen = xc + xm;
                y_cen = yc + ym;

                % Principal moments and orientation
                I = [ Iuu  -Iuv ;
                    -Iuv   Ivv ];
                [ eig_vec, eig_val ] = eig(I);
                I1 = eig_val(1,1); % principal MOI about 1 axis
                I2 = eig_val(2,2); % principal MOI about 2 axie
                ang1 = atan2( eig_vec(2,1), eig_vec(1,1) ); % orientation of 1 axis
                ang2 = atan2( eig_vec(2,2), eig_vec(1,2) ); % orientation of 2 axis
                
                % Plot the centoid output from Polygeom on last graph
                plot(x_cen,y_cen,'b+')
                
                % Preallocate memory to store the centroid for each slice
                % in a column vector and save the x and y coordinates
                
                centroid_x (j,:) = x_cen;
                centroid_y (j,:) = y_cen;
                
                % Section modulus is resistance to bending.  Here it is
                % Z=I/c where I is the centroidal MOI aboout the axis of
                % bending (the x axis) divided by the extreme fiber on the
                % failure surface (the medial surface is in tension) for a
                % tibia. A femur is tested about the mediolateral axis with
                % the anterior surface in tension
                
                if bone == 't'
                    section_mod = Iuu/medial_extreme;
                else
                    section_mod = Ivv/anterior_extreme;
                end
                
                %********************** SLICE OUTPUT***********************
                
                % For output purposes, we need the inner and outer fibers
                % converted to distance from pixels
                inner_fiber = inner_fiber * res;
                outer_fiber = outer_fiber * res;
                
                % Profiles can be dumbed into individual matirices which
                % will be added to duting this loop and then combined into
                % a single matrix after loop has ended for data output
                peri_out (j,:) = outer_fiber;
                endo_out (j,:) = inner_fiber;
                
                % Convert geometric props to proper units
                total_cs_area = total_cs_area * 1e-6;
                marrow_area = marrow_area * 1e-6;
                cortical_area = cortical_area * 1e-6;
                cort_thickness = cort_thickness * 1e-3;
                AP_width = AP_width * 1e-3;
                ML_width = ML_width * 1e-3;
                periosteal_BS = periosteal_BS * 1e-3;
                endocortical_BS = endocortical_BS * 1e-3;
                Iap = Iuu * 1e-12;
                Iml = Ivv * 1e-12;
                Imin = I1 * 1e-12;
                Imax = I2 * 1e-12;
                ang_min=ang1*180/pi;
                ang_max=ang2*180/pi;
                medial_extreme = medial_extreme * 1e-3;
                anterior_extreme = anterior_extreme * 1e-3;
                section_mod = section_mod * 1e-9;
                
                
                % Shift max and min angles into the 1st and 2nd Cartesian
                % quadrants
                if ang_min>=180
                    ang_min=ang_min-180;
                elseif ang_min<0
                    ang_min=ang_min+180;
                else
                    % do nothing
                end
                
                if ang_max>=180
                    ang_max=ang_max-180;
                elseif ang_max<0
                    ang_max=ang_max+180;
                else
                    % do nothing
                end
                
                % Store the output from each slice as a row in geom_out
                geometry = [total_cs_area marrow_area cortical_area cort_thickness ...
                    periosteal_BS endocortical_BS Imax ang_max Imin ang_min AP_width ML_width APtoMLratio  Iap ...
                    Iml section_mod medial_extreme anterior_extreme tmd_HA];
                geom_out(j,:) = geometry;
                
            end
            %********************** SAMPLE OUTPUT**************************
            
            % The matrice for angle, outer fiber, and inner fiber are
            % combined
            profiles (1,:) = 45:ang:405;
            profiles (2:length(slices)+1,:) = peri_out;
            profiles (length(slices)+2:2*length(slices)+1,:) = endo_out;
            profiles (2*length(slices)+2,:) = mean(profiles(2:length(slices)+1,:));
            profiles (2*length(slices)+3,:) = mean(profiles(length(slices)+2:2*length(slices)+1,:));
            
            % Overlay the average profile of all slices with the centroid
            % marked
            subplot(1, 2, 2)
            avg_peri = profiles (2*length(slices)+2,:);
            avg_endo = profiles (2*length(slices)+3,:);
            avg_peri_x = avg_peri.*cos(angle_rad);
            avg_peri_y = avg_peri.*sin(angle_rad);
            avg_endo_x = avg_endo.*cos(angle_rad);
            avg_endo_y = avg_endo.*sin(angle_rad);
            x_peri_min = abs(min(avg_peri_x));
            y_peri_min = abs(min(avg_peri_y));
            avg_peri_x = avg_peri_x+x_peri_min;
            avg_peri_x = fliplr(avg_peri_x);
            avg_peri_y = avg_peri_y+y_peri_min;
            avg_peri_y = fliplr(avg_peri_y);
            avg_endo_x = avg_endo_x+x_peri_min;
            avg_endo_y = avg_endo_y+y_peri_min;
            x_perimeter_avg = [avg_peri_x avg_endo_x];
            y_perimeter_avg = [avg_peri_y avg_endo_y];
            plot(x_perimeter_avg,y_perimeter_avg, 'r')
            hold on
            x_cen_avg = mean(centroid_x);
            y_cen_avg = mean(centroid_y);
            plot(x_cen_avg,y_cen_avg,'r+')
            axis equal
            axis tight
            xlabel('x-position in um')
            ylabel('y-position in um')
            title('Calculated Profiles')
            hold off
            
            % Prevent the polar subplot from shrinking (standard Matlab
            % error)
            h1 = subplot(1,2,1);  % save the handle of the subplot
            ax1=get(h1,'position'); % save the position as ax
            set(h1,'position',ax1);   % manually setting this holds the position
            
            % Create a polar plot of the profile with the Imax and Imin
            % axes
            subplot(1, 2, 1)
            polarplot(360,1000); % set  the axes for the polar plot
            hold on
            polarplot(angle_rad,avg_endo, '-b');
            polarplot(angle_rad,avg_peri, '-b');
            title('Polar Plot of Avg Profile')
            line_r = -max(avg_peri)+50:max(avg_peri)+50;
            ang_max_rad = mean(geom_out(1:length(slices),8)).*pi./180;
            ang_min_rad = mean(geom_out(1:length(slices),10)).*pi./180;
            ang_max_rad = ones(1,length(line_r)).*ang_max_rad;
            ang_min_rad = ones(1,length(line_r)).*ang_min_rad;
            pol_max = polarplot(ang_max_rad, line_r, '-r');
            pol_min = polarplot(ang_min_rad, line_r, '--r');
            legend([pol_max,pol_min],'\theta max','\theta min','Location','southoutside','Orientation','horizontal')
            hold off
            
            print ('-dpng', filename)
            
            % The matrice for angle, outer fiber and inner fiber are
            % shifted to begin at 0 degrees
            Ao = 315/ang+1;
            ang_shift = 0:ang:360;
            prof_shift1 = profiles(1:2*length(slices)+3,Ao:end);
            prof_shift2 = profiles(1:2*length(slices)+3,2:Ao);
            prof_shift = horzcat(prof_shift1, prof_shift2);
            prof_shift(1, :) = ang_shift;
            prof_cell=num2cell(prof_shift);
 
            % Create a cell array for prof_avg
            prof_mean_peri = prof_shift(2*length(slices)+2,:);
            prof_cell_peri = num2cell(prof_mean_peri);
            prof_peri_avg(m,:)=prof_mean_peri;
            prof_out_peri_cell(ppp, :) = prof_cell_peri;
            prof_mean_endo = prof_shift(2*length(slices)+3,:);
            prof_cell_endo = num2cell(prof_mean_endo);
            prof_out_endo_cell(ppp, :) = prof_cell_endo;
            prof_endo_avg(m,:)=prof_mean_endo;
            
            % Calculate average TMD using total counts from each slice
            % instead of a simple average as before because the pixel count
            % will change between slices
            tot_gs = sum(tmd_gs(:,2), 1); % total greyscale intensity for all slices
            tot_px = sum(tmd_gs(:,3), 1); % total pixels for all slices
            tmd_gs_avg = tot_gs/tot_px; % average greyscale intensity for all slices
            tmd_ac_avg = tmd_gs_avg.*ac_step; % convert to attenuation coefficient (AC)
            tmd_HA_avg = (tmd_ac_avg - eq_num)/eq_denom; % convert AC to BMD using equation
            
            % Create a cell array for geom_avg
            geom_mean = mean(geom_out(1:length(slices),:));
            mean_cell=num2cell(geom_mean);
            geom_out_cell(ppp, 1:18) = mean_cell(1: 18);
            geom_out_cell(ppp, 19) = num2cell(tmd_HA_avg);
            cd('..')
            ppp=ppp+1;
            
%             timer = toc;
%             fprintf('\n Sample %s took %u seconds.',filename,timer)
        end
    end
    specimen_count(k)=count;
    data_1=mean(prof_peri_avg);
    data_2=mean(prof_endo_avg);
    avg_prof{k,:}={name, data_1, data_2};
end       
    % Trim the variables to avoid dimension errors
    sample_list(:,all(cellfun(@isempty,sample_list),1)) = [];
    %geom_out_cell(all(cellfun(@isempty,geom_out_cell),2),:) = [];
    partial = length(sample_list);
    geom_out_cell(partial+1:end,:) = [];
    prof_out_peri_cell(partial+1:end,:) = [];
    prof_out_endo_cell(partial+1:end,:) = [];
    
    prof_data=[avg_prof{1,1}{1,2}; avg_prof{1,1}{1,3}; ...
    avg_prof{2,1}{1,2}; avg_prof{2,1}{1,3}; ...
    avg_prof{3,1}{1,2}; avg_prof{3,1}{1,3}; ...
    avg_prof{4,1}{1,2}; avg_prof{4,1}{1,3}];

%% Mean property and profile output for each bone

% Create cell array for output containing all column and row headers along
% with the profiles and save the data to an xlsx spreadsheet.
peri_cell = ['Periosteal'; sample_list'];
endo_cell = ['Endocortical'; sample_list'];
col_prof_cell = [peri_cell; ' '; endo_cell];
blank = cell(1, 360/ang +1);
prof_out = [prof_cell(1, :); prof_out_peri_cell; blank; prof_cell(1, :); prof_out_endo_cell];
prof_mean_out = horzcat(col_prof_cell, prof_out);
xlswrite('CTprof_avg_ALL.xlsx', prof_mean_out, 'Raw Data', 'A1')

% Create summary summary prof file
row_1={A_name, ' ', B_name, ' ', C_name, ' ', D_name,};
row_2={'Periosteal','Endocortical','Periosteal','Endocortical','Periosteal','Endocortical','Periosteal','Endocortical'};

xlswrite('CTprof_AVG',row_1,1,'a1')
xlswrite('CTprof_AVG',row_2,1,'a2')
xlswrite('CTprof_AVG',prof_data',1,'a3:h723')

% Create cell array for output containing all column and row headers along
% with the geometric properties and save the data to an xlsx spreadsheet.
geom_out_data=cell2mat(geom_out_cell);
geom_out_cell = horzcat(sample_list', group_list', geom_out_cell); % label the rows
headers = {'Total CSA (mm^2)', 'Marrow Area (mm^2)', 'Cortical Area(mm^2)', 'Cortical Thickness (mm)', 'Periosteal BS (mm)', 'Endocortical BS (mm)',  'Imax (mm^4)', 'Theta max (deg)', 'Imin (mm^4)','Theta min (deg)', 'AP Width (mm)', 'ML Width (mm)', 'AP/ML',  'Iap (mm^4)', 'Iml (mm^4)',   'Section Modulus (mm^3)', 'Medial Extreme (mm)', 'Anterior Extreme (mm)','TMD (g/cm^3 HA)'};
headers = horzcat(' ', 'Group', headers);
geom_mean_out = [headers; geom_out_cell]; % add column titles
xlswrite('CTgeom_avg.xlsx', geom_mean_out, 'Raw Data', 'A1')

%% Create folder with .csv file for each output property
mkdir('Data Files');
cd('Data Files');
max_s=max(specimen_count);

org_rows={cat1_1; cat1_2};
org_col=horzcat(' ', cat2_1, num2cell(zeros(1,max_s-1)), cat2_2, num2cell(zeros(1,max_s-1)));

% Get indexes for each group set
ca1=1;
ca2=specimen_count(1);
cb1=ca2+1;
cb2=cb1+specimen_count(2)-1;
cc1=cb2+1;
cc2=cc1+ specimen_count(3)-1;
cd1=cc2+1;
cd2=cd1+specimen_count(4)-1;

data_a=geom_out_data(ca1:ca2,:);
data_a(min(size(data_a))+1:max_s,:) = 0;
data_b=geom_out_data(cb1:cb2,:);
data_b(min(size(data_b))+1:max_s,:) = 0;
data_c=geom_out_data(cc1:cc2,:);
data_c(min(size(data_c))+1:max_s,:) = 0;
data_d=geom_out_data(cd1:cd2,:);
data_d(min(size(data_d))+1:max_s,:) = 0;

% Total CSA
data_print=num2cell([data_a(:,1), data_b(:,1); data_c(:,1) data_d(:,1)]');
data_print=[org_rows, data_print];
data_out=[org_col; data_print];
writecell(data_out,'Total_CSA.csv')

% Marrow Area
data_print=num2cell([data_a(:,2), data_b(:,2); data_c(:,2) data_d(:,2)]');
data_print=[org_rows, data_print];
data_out=[org_col; data_print];
writecell(data_out,'Marrow_Area.csv')

% Cortical Area
data_print=num2cell([data_a(:,3), data_b(:,3); data_c(:,3) data_d(:,3)]');
data_print=[org_rows, data_print];
data_out=[org_col; data_print];
writecell(data_out,'Cortical_Area.csv')

% Cortical Thickness
data_print=num2cell([data_a(:,4), data_b(:,4); data_c(:,4) data_d(:,4)]');
data_print=[org_rows, data_print];
data_out=[org_col; data_print];
writecell(data_out,'Cortical_Thickness.csv')

% Periosteal BS
data_print=num2cell([data_a(:,5), data_b(:,5); data_c(:,5) data_d(:,5)]');
data_print=[org_rows, data_print];
data_out=[org_col; data_print];
writecell(data_out,'Periosteal_BS.csv')

% Endocortical BS
data_print=num2cell([data_a(:,6), data_b(:,6); data_c(:,6) data_d(:,6)]');
data_print=[org_rows, data_print];
data_out=[org_col; data_print];
writecell(data_out,'Endocortical_BS.csv')

% Imax
data_print=num2cell([data_a(:,7), data_b(:,7); data_c(:,7) data_d(:,7)]');
data_print=[org_rows, data_print];
data_out=[org_col; data_print];
writecell(data_out,'Imax.csv')

% Imin
data_print=num2cell([data_a(:,9), data_b(:,9); data_c(:,9) data_d(:,9)]');
data_print=[org_rows, data_print];
data_out=[org_col; data_print];
writecell(data_out,'Imin.csv')

% Section Modulus 
data_print=num2cell([data_a(:,16), data_b(:,16); data_c(:,16) data_d(:,16)]');
data_print=[org_rows, data_print];
data_out=[org_col; data_print];
writecell(data_out,'Section_Modulus.csv')

% TMD
data_print=num2cell([data_a(:,19), data_b(:,19); data_c(:,19) data_d(:,19)]');
data_print=[org_rows, data_print];
data_out=[org_col; data_print];
writecell(data_out,'TMD.csv')

% tot_timer = toc(tot);
% tot_msg = ['\n\n Total time for all samples was ' num2str(tot_timer) ' seconds.'];
% avg_msg = [' Average time per sample was ' num2str(tot_timer/length(key)) ' seconds. \n\n'];
% fprintf(tot_msg)
% fprintf(avg_msg)

end
