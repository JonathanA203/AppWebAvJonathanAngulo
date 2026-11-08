using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV3
{
    public partial class Song : System.Web.UI.Page
    {
        string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;
        protected void Page_Load(object sender, EventArgs e)
        {
            EntityLayer.PoliMusicV3.User user = Session[EntityLayer.PoliMusicV3.Util.Constants.USER] as EntityLayer.PoliMusicV3.User;
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
            gvSongs.DataSource = new BLL.PoliMusicV3.Song(strConnString).Read();
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
            EntityLayer.PoliMusicV3.Song song = new EntityLayer.PoliMusicV3.Song();
            song.Name = songName;
            string strSongFolder = ConfigurationManager.AppSettings["songFolderPath"];
            string newFileName = songName.Replace(" ", string.Empty);
            song.Path = strSongFolder + newFileName + System.IO.Path.GetExtension(fupSong.FileName).ToLower();

            string messageValidations = new BLL.PoliMusicV3.Song(strConnString).ValidationsDuplicated(song);
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
            new BLL.PoliMusicV3.Song(strConnString).Save(song);
            txtSongName.Text = string.Empty;
            lblMessage.Text = "Song uploaded successfully!";
            LoadSongs();
        }
    }

}
