using System;
using System.Collections.Generic;
using System.Configuration;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV2
{
    public partial class Song : System.Web.UI.Page
    {
        string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;
        protected void Page_Load(object sender, EventArgs e)
        {
            Entities.User user = Session[Util.Constants.USER] as Entities.User;
            if (user == null)
            {
                Response.Redirect("Login.aspx", true);
                return;
            }
            if (!Page.IsPostBack)
            {
                lblUserName.Text = Server.HtmlEncode(user.UserName);
                lblMessage.Text = string.Empty;
                LoadSongs();
            }
        }

        private void LoadSongs()
        {
            gvSongs.DataSource = new BLL.Song(strConnString).Read();
            gvSongs.DataBind();
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            string songName = txtSongName.Text.Trim();
            if (songName == string.Empty)
            {
                lblMessage.Text = "Song name is required!";
                return;
            }
            Entities.Song song = new Entities.Song();
            song.Name = songName;
            string strSongFolder = ConfigurationManager.AppSettings["songFolderPath"];
            string newFileName = songName.Replace(" ", string.Empty);
            song.Path = strSongFolder + newFileName + System.IO.Path.GetExtension(fupSong.FileName).ToLower();

            string messageValidations = new BLL.Song(strConnString).ValidationsDuplicated(song);
            if (messageValidations != string.Empty)
            {
                lblMessage.Text = messageValidations;
                return;
            }
            messageValidations = new FileManagement().SaveSongOnServer(fupSong, strSongFolder, newFileName);
            if (messageValidations != string.Empty)
            {
                lblMessage.Text = messageValidations;
                return;
            }
            new BLL.Song(strConnString).Save(song);
            txtSongName.Text = string.Empty;
            lblMessage.Text = "Song uploaded successfully!";
            LoadSongs();
        }
    }

}
