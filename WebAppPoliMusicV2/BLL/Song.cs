using System;
using System.Collections.Generic;
using System.Data;
using System.Web;

namespace WebAppPoliMusicV2.BLL
{
    public class Song
    {
        private string strConnectionString;
        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }
        public int Save(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Save(song);
        }
        public DataTable Read()
        {
            return new DAL.Song(strConnectionString).Read();
        }
        public int Update(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Update(song);
        }
        public int Delete(int id)
        {
            return new DAL.Song(strConnectionString).Delete(id);
        }
        public int CheckExistSong(string songName)
        {
            return new DAL.Song(strConnectionString).CheckExistSong(songName);
        }
        public string ValidationsDuplicated(Entities.Song song)
        {
            int returnValue = CheckExistSong(song.Name);
            if (returnValue != 0)
            {
                return "Song name already registered!";
            }
            return string.Empty;
        }
    }

}
