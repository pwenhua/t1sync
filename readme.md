# T1Sync C# — Visual Studio Setup

NuGet install ClosedXML

#
the replace logic is: like for column name 'Near Power Line' , all meta data matters, it is located under 'asset_type' level 2, locate node having "AttributeCode": "ASSET_TYPE" and "SearchPath": "Tree\Street Tree" , split the SearchPath by '\' we get two string - indicate this is level 2 , level 1 has no '\' at all, by 'AttributeCode' and 'SearchPath' we should only get one node ; meta data says it is 'Userfield1', then under the node, replace value of 'AttributeItemUserfield1'with the cell value