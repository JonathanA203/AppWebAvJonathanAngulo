<%@ Page Title="Song List" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV3.Song" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <main aria-labelledby="title">
        <h2 id="title">Welcome, <asp:Label ID="lblUserName" runat="server"></asp:Label></h2>

        <h3>Available songs</h3>
        <asp:GridView ID="gvSongs" runat="server" AutoGenerateColumns="False" DataKeyNames="ID_SONG"
            CssClass="table table-striped" GridLines="None"
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

        <fieldset class="mb-3">
            <legend>Add a song (MP3)</legend>
            <div class="mb-2">
                <label for="<%= txtSongName.ClientID %>">Song name*:</label><br />
                <asp:TextBox ID="txtSongName" runat="server" CssClass="form-control"></asp:TextBox>
            </div>
            <div class="mb-2">
                <label for="<%= fupSong.ClientID %>">File*:</label><br />
                <asp:FileUpload ID="fupSong" runat="server" CssClass="form-control" />
            </div>
            <asp:Button ID="btnUpload" runat="server" Text="Upload" CssClass="btn btn-primary" OnClick="btnUpload_Click" />
        </fieldset>
        <asp:Label ID="lblMessage" ForeColor="Red" Font-Bold="true" runat="server"></asp:Label>
    </main>
</asp:Content>
