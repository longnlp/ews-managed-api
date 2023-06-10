namespace Test
{
    using Microsoft.Exchange.WebServices.Data;

    public class EXOArchiveMailboxTest : BaseTokenTest
    {
        [Fact]
        public void TestMailboxSize()
        {
            var exo = CreateEXOServiceWithAppCredentialsByMSAL(m365Context.AppAccount);

            var PidTagMessageSizeExtended = new ExtendedPropertyDefinition(0xe08, MapiPropertyType.Long);

            var PidTagNormalMessageSizeExtended = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.Common, "PR_NORMAL_MESSAGE_SIZE_EXTENDED", MapiPropertyType.Double);


            var folderView = new FolderView(100);
            folderView.PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, PidTagMessageSizeExtended, PidTagNormalMessageSizeExtended);
            folderView.Traversal = FolderTraversal.Deep;


            FindFoldersResults folders = null;

            var items = new Dictionary<string, Tuple<long, long>>();
            var folderName = new Dictionary<string, string>();
            long totalSize = 0;


            while (folders == null || folders.MoreAvailable)
            {
                //output folders
                folderView.Offset = folders == null ? 0 : folders.NextPageOffset.Value;
                folders = exo.FindFolders(WellKnownFolderName.ArchiveMsgFolderRoot, folderView);

                foreach (var folder in folders)
                {
                    string parentName = null;
                    string fullName = folder.DisplayName;
                    if(folderName.TryGetValue(folder.ParentFolderId.ToString(), out parentName))
                    {
                        fullName = parentName + "/" + folder.DisplayName;
                    }

                    long folderSize;

                    folder.TryGetProperty(PidTagMessageSizeExtended, out folderSize);
                    totalSize += folderSize;

                    items.Add(fullName, new Tuple<long, long>(folderSize, folder.TotalCount));
                    folderName.Add(folder.Id.ToString(), fullName);
                }
            }

            Assert.True(totalSize > 0);
            
        }
    }
}