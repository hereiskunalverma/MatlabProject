clc                  % Clear command window.
clear all            %Clear variables and functions from memory
close all            % closes all the open figure windows
home                 % Send the cursor home

database = 'database.xls' % Reading data from the excel file
[numbers,text] = xlsread(database) % reading numbers and text


len=length(text) % number of certificates to be generated

% Converting image to .tif format
img=imread("for completing courses in.png"); % Reading image
imwrite(img,"CerTemplate.tif","tif");        % Saving image as .tif format

%Itrain = double(imread('CerTemplate.tif'));
% Extract the individual red, green, and blue color channels.
%redChannel = Itrain(:, :, 1);
%greenChannel = Itrain(:, :, 2);
%blueChannel = Itrain(:, :, 3);
% Channel = Itrain(:, :, 4); % We don't need 4th channel so skip it
% % Recombine separate color channels into an RGB image.
%blankimage = uint8(cat(3,redChannel, greenChannel, blueChannel ));
blankimage = imread("CerTemplate.tif");
% Read Certificate Template

for i=1:len
    for j= 2:2 
        names(i,j)=text(i,j)
    end
end
% Obtain names from the text variable in 2nd column

for i=1:len
    for j= 3:3
      courses(i,j)=text(i,j)
    end
end
% Obtain courses names from 3rd column

% No need of first column 
for i=2:len
        text_all=[names(i,2) courses(i,3)]
        % combine names and courses
        
        position=[3569 2599; 3633 3515];
        % Can obtain position using MSPaint and get the pixels value
        
        RGB = insertTxet(blankimage,position,text_all,'FontSize',100,'BoxOpacity',0);
        %Provide parameters for function insertText
        %Font size is 100 and opacity of box is 0 means fully transparent
        
        figure
        imshow(RGB)        
        %Obtain and display figure in color
        
        y=i-1
        file=['Certificate_' num2str(y)]
        result=[file '.png']
        saveas(gcf,result)
        % generated images have .png extension (change as per convenience)
end    

