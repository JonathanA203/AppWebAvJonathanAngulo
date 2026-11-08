using System;
using System.Collections.Generic;
using System.Data;

namespace BLL.PoliMusicV3
{
    public class Song
    {
        private string strConnectionString;
        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }
        public int Save(EntityLayer.PoliMusicV3.Song song)
        {
            return new DAL.PoliMusicV3.Song(strConnectionString).Save(song);
        }
        public DataTable Read()
        {
            return new DAL.PoliMusicV3.Song(strConnectionString).Read();
        }
        public int Update(EntityLayer.PoliMusicV3.Song song)
        {
            return new DAL.PoliMusicV3.Song(strConnectionString).Update(song);
        }
        public int Delete(int id)
        {
            return new DAL.PoliMusicV3.Song(strConnectionString).Delete(id);
        }
        public int CheckExistSong(string songName)
        {
            return new DAL.PoliMusicV3.Song(strConnectionString).CheckExistSong(songName);
        }
        public string ValidationsDuplicated(EntityLayer.PoliMusicV3.Song song)
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
