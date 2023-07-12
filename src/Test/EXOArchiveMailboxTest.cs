namespace Test
{
    using Microsoft.Exchange.WebServices.Data;
    using Xunit.Abstractions;

    public class EXOArchiveMailboxTest : BaseTokenTest
    {
        public EXOArchiveMailboxTest(ITestOutputHelper output) : base(output)
        {
        }

        [Fact]
        public void TestMailboxSize()
        {
            var exo = CreateEXOServiceWithAppCredentialsByMSAL(m365Context.AppAccount);

            var folders = exo.LoadAllFolders(WellKnownFolderName.ArchiveMsgFolderRoot);
            output.WriteLine("{0, -55}{1, -15}", "Name", "Folder Size");
            foreach (var folder in folders.Values)
            {
                output.WriteLine("{0, -55}{1, -15}", folder.FullName, folder.FolderSize);
            }

            Assert.True(folders.Count > 0);
            
        }
    }
}