<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>Song List</title>
</head>
<body>
    <form id="formSong" runat="server">
        <h2>Welcome, <asp:Label ID="lblUserName" runat="server"></asp:Label></h2>
        <asp:HyperLink ID="lnkLogout" NavigateUrl="~/Logout.aspx" runat="server">Logout</asp:HyperLink>

        <h3>Available songs</h3>
        <asp:GridView ID="gvSongs" runat="server" AutoGenerateColumns="False" DataKeyNames="ID_SONG"
            EmptyDataText="There are no songs available.">
            <Columns>
                <asp:BoundField DataField="SONG_NAME" HeaderText="Song" />
                <asp:BoundField DataField="ARTISTS" HeaderText="Artists" />
                <asp:BoundField DataField="ALBUM_NAME" HeaderText="Album" />
                <asp:BoundField DataField="PLAYS" HeaderText="Plays" />
                <asp:TemplateField HeaderText="Play">
                    <ItemTemplate>
                        <audio controls="controls" preload="none" src='<%# ResolveUrl(Eval("SONG_PATH").ToString()) %>'></audio>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>

        <fieldset>
            <legend>Add a song (MP3)</legend>
            Song name*:<br /><asp:TextBox ID="txtSongName" runat="server"></asp:TextBox><br />
            File*:<br /><asp:FileUpload ID="fupSong" runat="server" /><br />
            <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click" />
        </fieldset>
        <asp:Label ID="lblMessage" ForeColor="Red" Font-Bold="true" runat="server"></asp:Label>
    </form>
</body>
</html>
