function theStruct = parseXML(fullfile)
% PARSEXML Convert XML file to a MATLAB structure. Includes capability to
% filter invalid XML syntax out of the file and write repaired file.
%
% Original Code taken from Matlab Help (see doc xmlwrite)
% 
% Modified to handle the case of a not well-formed XML tree. This occurs in
% older cases of WRII v10.16 (and maybe others). This function will write a
% NEW MMT file that has been filtered. The new file will be in the same
% folder of the original file, and will have the suffix "_repaired".
% 
% Written by: Frank L. Engel, USGS IL WSC

% See if the MMT is a well-formed XML doc. If a java exception is thrown,
% go ahead and fix the bad characters
try
    tree = xmlread(fullfile);

% Ok, that didn't work. Throw a warning message, open the file manually and
% repair it. Write a new version in the current working directory.
catch err_bad_xml_markup
    [pathstr, filename, ext] = fileparts(fullfile);
    warning([...
        'Failed to read XML file: %s\n'...
        '   in directory: %s\n'...
        'Attempting to repair file and continue...'],filename,pathstr);
    
    % Open the file manually, and read all text into a char array (this is
    % the UTF-8 decimals for each char type
    fid = fopen(fullfile, 'rb');
    str = fread(fid, [1, inf], 'char');
    
    % Create a logical array of the valid data == TRUE
    idx_keep = (~isstrprop(str,'cntrl')) | (str==10) | (str==13);
    
    % Delete any invalid characters    
    str(~idx_keep) = '';
    
    % Re-encode the string in UTF-8 and close the file
    encoded_str = unicode2native(char(str), 'UTF-8');
    fclose(fid);
    
    % Create and open a new file with suffix "_repaired". If the file
    % already exists, it will overwrite without warning.
    sep = filesep;
    repairedfile = [pathstr sep filename '_repaired' ext];
    fid = fopen(repairedfile, 'wt');
    fwrite(fid, encoded_str, 'uint8');
    fclose(fid);
    
    % Now, the new file should be good, go ahead and try opening it as XML
    % again.
    try
    tree = xmlread(repairedfile);
    disp('Repair successful.')
    catch err_badxml_secondtry
        error([...
        'Failed to read repaired XML file: %s\n'...
        '   in directory: %s\n'],[filename '_repaired' ext],pathstr);
    end
end

% Recurse over child nodes. This could run into problems 
% with very deeply nested trees.
try
   theStruct = parseChildNodes(tree);
catch
   error('Unable to parse XML file %s.',fullfile);
end


% ----- Subfunction PARSECHILDNODES -----
function children = parseChildNodes(theNode)
% Recurse over node children.
children = [];
if theNode.hasChildNodes
   childNodes = theNode.getChildNodes;
   numChildNodes = childNodes.getLength;
   allocCell = cell(1, numChildNodes);

   children = struct(             ...
      'Name', allocCell, 'Attributes', allocCell,    ...
      'Data', allocCell, 'Children', allocCell);

    for count = 1:numChildNodes
        theChild = childNodes.item(count-1);
        children(count) = makeStructFromNode(theChild);
    end
end

% ----- Subfunction MAKESTRUCTFROMNODE -----
function nodeStruct = makeStructFromNode(theNode)
% Create structure of node info.

nodeStruct = struct(                        ...
   'Name', char(theNode.getNodeName),       ...
   'Attributes', parseAttributes(theNode),  ...
   'Data', '',                              ...
   'Children', parseChildNodes(theNode));

if any(strcmp(methods(theNode), 'getData'))
   nodeStruct.Data = char(theNode.getData); 
else
   nodeStruct.Data = '';
end

% ----- Subfunction PARSEATTRIBUTES -----
function attributes = parseAttributes(theNode)
% Create attributes structure.

attributes = [];
if theNode.hasAttributes
   theAttributes = theNode.getAttributes;
   numAttributes = theAttributes.getLength;
   allocCell = cell(1, numAttributes);
   attributes = struct('Name', allocCell, 'Value', ...
                       allocCell);

   for count = 1:numAttributes
      attrib = theAttributes.item(count-1);
      attributes(count).Name = char(attrib.getName);
      attributes(count).Value = char(attrib.getValue);
   end
end