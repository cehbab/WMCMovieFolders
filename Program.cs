using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Authentication;
using System.Text;
using System.Xml.Serialization;

namespace WMCMovieFolders
{

    [DataContract]
    public class Page
    {
        public Page() { }

        [DataMember]
        public int page { get; set; }

        [DataMember]
        public Result[] results { get; set; }

        [DataMember]
        public int total_pages { get; set; }

        [DataMember]
        public int total_results { get; set; }
    }

    [DataContract]
    public class Result
    {
        public Result() { }

        [DataMember]
        public string id { get; set; }

        [DataMember]
        public int[] genre_ids { get; set; }

        [DataMember]
        public string original_language { get; set; }

        [DataMember]
        public string original_title { get; set; }

        [DataMember]
        public string overview { get; set; }

        [DataMember]
        public float popularity { get; set; }

        [DataMember]
        public string poster_path { get; set; }

        [DataMember]
        public string release_date { get; set; }

        [DataMember]
        public string title { get; set; }

        [DataMember]
        public bool video { get; set; }

        [DataMember]
        public float vote_average { get; set; }

        [DataMember]
        public int vote_count { get; set; }
    }

    [DataContract]
    public class Movie
    {
        public Movie() { }

        [DataMember]
        public bool adult { get; set; }

        [DataMember]
        public string backdrop_path { get; set; }

        [DataMember]
        public long budget { get; set; }

        [DataMember]
        public Genre[] genres { get; set; }

        [DataMember]
        public string id { get; set; }

        [DataMember]
        public string imdb_id { get; set; }

        [DataMember]
        public string original_language { get; set; }

        [DataMember]
        public string original_title { get; set; }

        [DataMember]
        public string overview { get; set; }

        [DataMember]
        public float popularity { get; set; }

        [DataMember]
        public string poster_path { get; set; }

        [DataMember]
        public Studio[] production_companies { get; set; }

        [DataMember]
        public string release_date { get; set; }

        [DataMember]
        public long revenue { get; set; }

        [DataMember]
        public int runtime { get; set; }

        [DataMember]
        public string status { get; set; }

        [DataMember]
        public string tagline { get; set; }

        [DataMember]
        public string title { get; set; }

        [DataMember]
        public bool video { get; set; }

        [DataMember]
        public float vote_average { get; set; }

        [DataMember]
        public int vote_count { get; set; }

        [DataMember]
        public Premieres release_dates { get; set; }

        [DataMember]
        public Credits credits { get; set; }
    }

    [DataContract]
    public class Genre
    {
        public Genre() { }

        [DataMember]
        public int id { get; set; }

        [DataMember]
        public string name { get; set; }
    }

    [DataContract]
    public class Studio
    {
        public Studio() { }

        [DataMember]
        public int id { get; set; }

        [DataMember]
        public string logo_path { get; set; }

        [DataMember]
        public string name { get; set; }

        [DataMember]
        public string origin_country { get; set; }
    }


    [DataContract]
    public class Premieres
    {
        public Premieres() { }

        [DataMember]
        public Releases[] results { get; set; }
    }

    [DataContract]
    public class Releases
    {
        public Releases() { }

        [DataMember]
        public string iso_3166_1 { get; set; }

        [DataMember]
        public Release[] release_dates { get; set; }
    }

    [DataContract]
    public class Release
    {
        public Release() { }

        [DataMember]
        public string certification { get; set; }

        [DataMember]
        public string iso_639_1 { get; set; }

        [DataMember]
        public string note { get; set; }

        [DataMember]
        public string release_date { get; set; }

        [DataMember]
        public Showing type { get; set; }
    }

    public enum Showing
    {
        Premiere = 1,
        Limited = 2,
        Theatrical = 3,
        Digital = 4,
        Physical = 5,
        TV = 6
    }


    [DataContract]
    public class Credits
    {
        public Credits() { }

        [DataMember]
        public Cast[] cast { get; set; }

        [DataMember]
        public Crew[] crew { get; set; }
    }


    [DataContract]
    public class Person
    {
        public Person() { }

        [DataMember]
        public bool adult { get; set; }

        [DataMember]
        public Gender gender { get; set; }

        [DataMember]
        public int id { get; set; }

        [DataMember]
        public string known_for_department { get; set; }

        [DataMember]
        public string name { get; set; }

        [DataMember]
        public string original_name { get; set; }

        [DataMember]
        public float popularity { get; set; }

        [DataMember]
        public string profile_path { get; set; }

        [DataMember]
        public string credit_id { get; set; }
    }

    public enum Gender
    {
        None = 0,
        Female = 1,
        Male = 2,
        NonBinary = 3
    }


    [DataContract]
    public class Cast : Person
    {
        public Cast() { }

        [DataMember]
        public int cast_id { get; set; }

        [DataMember]
        public string character { get; set; }

        [DataMember]
        public int order { get; set; }
    }

    [DataContract]
    public class Crew : Person
    {
        public Crew() { }

        [DataMember]
        public string department { get; set; }

        [DataMember]
        public string job { get; set; }
    }


    [XmlRoot("Disc")]
    public class DVDID
    {
        public static DVDID Create(string wmcid, string name)
        {
            DVDID obj = new DVDID();
            obj.name = name;
            obj.id = wmcid;
            return obj;
        }

        public DVDID() { }

        [XmlElement(ElementName = "Name")]
        public string name { get; set; }

        [XmlElement(ElementName = "ID")]
        public string id { get; set; }
    }

    [XmlRoot("METADATA")]
    public class WMCID
    {
        public static WMCID Create(string wmcid, MDRDVD mdrdvd)
        {
            WMCID obj = new WMCID();
            obj.mdrdvd = mdrdvd;
            obj.DvdId = wmcid;
            return obj;
        }

        public WMCID() { }

        [XmlElement(ElementName = "MDR-DVD")]
        public MDRDVD mdrdvd { get; set; }

        public string DvdId { get; set; }
    }

    [XmlRoot("MDR-DVD")]
    public class MDRDVD
    {
        public class Title
        {
            public static Title Create(MDRDVD mdrdvd, Movie movie)
            {
                Title obj = new Title();
                obj.titleNum = 1;
                obj.titleTitle = mdrdvd.dvdTitle;
                obj.studio = mdrdvd.studio;
                obj.leadPerformer = mdrdvd.leadPerformer;
                obj.director = mdrdvd.director;
                obj.MPAARating = mdrdvd.MPAARating;
                obj.genre = mdrdvd.genre;
                obj.synopsis = movie.overview;
                return obj;
            }

            public Title() { }

            public int titleNum { get; set; }

            public string titleTitle { get; set; }

            public string studio { get; set; }

            public string leadPerformer { get; set; }

            public string director { get; set; }

            public string MPAARating { get; set; }

            public string genre { get; set; }
            
            public string synopsis { get; set; }

        }

        public class Rating
        {
            public string country;

            public string[] g;

            public string[] pg;

            public string[] m;

            public string[] m15;

            public string[] r18;
        }

        //Country: 'AU'
        //  Ratings: 'G'; 'PG'; 'M'; 'MA15+'; 'MA 15+'; 'R18+'; 'R 18+';
        //Country: 'US'
        //  Ratings: 'G'; 'PG'; 'PG-13'; 'R'; 'NC-17'; 'NR';
        //Country: 'GB'
        //  Ratings: 'U'; 'PG'; '12A'; '12'; '15'; '18'; 'R18';

        public static Rating[] ratings = new Rating[]
        {
            new Rating
            {
                country = "AU",
                g = new string[] { "G" },
                pg = new string[] { "PG" },
                m = new string[] { "M" },
                m15 = new string[] { "MA15+", "MA 15+" },
                r18 = new string[] { "R18+", "R 18+" }
            },
            new Rating
            {
                country = "US",
                g = new string[] { "G" },
                pg = new string[] { "PG" },
                m = new string[] { "PG-13" },
                m15 = new string[] { "R" },
                r18 = new string[] { "NC-17" }
                // "NR" is left blank, for No Rating
            },
            new Rating
            {
                country = "GB",
                g = new string[] { "U" },
                pg = new string[] { "PG" },
                m = new string[] { "12A", "12" },
                m15 = new string[] { "15" },
                r18 = new string[] { "18", "R18" }
            }
        };

        public static Rating AU = ratings[0];

        //public static string FindMPAARating(string country, string rating)
        //{
        //}
        //{
        //    //if (country.Equals("AU", StringComparison.CurrentCultureIgnoreCase)
        //    //{
        //    //    if (
        //    ////  Ratings: 'G'; 'PG'; 'M'; 'MA 15+'; 'MA15+'; 'R 18+'; 'R18+';
        //    //}
        //    //Country: 'US'
        //    //  Ratings: 'G'; 'PG'; 'PG-13'; 'R'; 'NC-17'; 'NR';
        //    //Country: 'GB'
        //    //  Ratings: 'U'; 'PG'; '12A'; '12'; '15'; '18'; 'R18';
        //}

        public static string FindRating(Movie movie)
        {
            foreach (Rating rating in ratings)
                foreach (Releases releases in movie.release_dates.results)
                    if (releases.iso_3166_1.Equals(rating.country, StringComparison.CurrentCultureIgnoreCase))
                        foreach (Release release in releases.release_dates)
                            if (rating.g.Contains(release.certification))
                                return AU.g[0];
                            else if (rating.pg.Contains(release.certification))
                                return AU.pg[0];
                            else if (rating.m.Contains(release.certification))
                                return AU.m[0];
                            else if (rating.m15.Contains(release.certification))
                                return AU.m15[0];
                            else if (rating.r18.Contains(release.certification))
                                return AU.r18[0];
            return "";
        }


        public static MDRDVD Create(Movie movie, string[] imgnames)
        {
            bool empty;

            bool pg = false;
            StringBuilder genres = new StringBuilder();
            empty = true;
            foreach (Genre genre in movie.genres)
            {
                if (!empty)
                    genres.Append("; ");
                genres.Append(genre.name);
                empty = false;

                pg = pg || genre.name.Equals("Family", StringComparison.CurrentCultureIgnoreCase);
            }

            string rating = FindRating(movie);

            StringBuilder studios = new StringBuilder();
            empty = true;
            foreach (Studio studio in movie.production_companies)
            {
                if (!empty)
                    studios.Append("; ");
                studios.Append(studio.name);
                empty = false;
            }

            string performers = (movie.credits.cast.Length != 0 ? movie.credits.cast[0].name : null);
            if (String.IsNullOrEmpty(performers))
                performers = "";

            StringBuilder directors = new StringBuilder();
            empty = true;
            foreach (Crew member in movie.credits.crew)
                if (member.job.Equals("Director", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (!empty)
                        studios.Append("; ");
                    directors.Append(member.name);
                    empty = false;
                }


            MDRDVD obj = new MDRDVD();
            obj.version = "5.0";
            obj.dvdTitle = movie.title;
            obj.studio = studios.ToString();
            obj.leadPerformer = performers;
            obj.director = directors.ToString();
            obj.MPAARating = rating;
            obj.releaseDate = movie.release_date.Replace('-', ' ');
            obj.genre = genres.ToString();
            obj.smallCoverParams = imgnames[0];
            obj.largeCoverParams = imgnames[1];
            obj.dataProvider = "AMG";
            obj.duration = movie.runtime;
            obj.title = Title.Create(obj, movie);
            return obj;
        }

        public MDRDVD() { }

        public string version { get; set; }

        public string dvdTitle { get; set; }

        public string studio { get; set; }

        public string leadPerformer { get; set; }

        public string director { get; set; }

        public string MPAARating { get; set; }

        public string releaseDate { get; set; }

        public string genre { get; set; }

        public string largeCoverParams { get; set; }

        public string smallCoverParams { get; set; }

        public string dataProvider { get; set; }

        public int duration { get; set; }

        public Title title { get; set; }

    }

    class Program
    {
        public const SslProtocols _Tls12 = (SslProtocols)0x00000C00;
        public const SecurityProtocolType Tls12 = (SecurityProtocolType)_Tls12;

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        static extern bool CreateHardLink(string lpFileName, string lpExistingFileName, IntPtr mustBeNull);

        //static string APIKEY = "eyJhbGciOiJIUzI1NiJ9.eyJhdWQiOiJlOTE1NDE0NGVkNTVmZTE3MjFhNGQyNDY3ZDQ0N2NkOSIsInN1YiI6IjY1YjU4ODU0NThlZmQzMDE3Y2NhYjY0ZiIsInNjb3BlcyI6WyJhcGlfcmVhZCJdLCJ2ZXJzaW9uIjoxfQ.zGVfNRR4j_uHhW4_KY6ViAfQSKNqB7RGziDwXjXpvLU";
        //static string APIKEY = "e9154144ed55fe1721a4d2467d447cd9";

        static string IMGWIDTH_SMALL = "w185";
        static string IMGWIDTH_LARGE = "w300";
        //static string IMGWIDTH = "w500";
        //static string IMGWIDTH = "w780";
        //static string IMGWIDTH = "original";


        static T FromJSON<T>(string json)
        {
            T obj = Activator.CreateInstance<T>();
            using (MemoryStream ms = new MemoryStream(Encoding.Unicode.GetBytes(json)))
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(obj.GetType());
                obj = (T)serializer.ReadObject(ms);
            }

            return obj;
        }


        static bool FetchData(string url, string filePath)
        {
            WebClient client = new WebClient();
            client.Headers[HttpRequestHeader.ContentType] = "application/binary";
            //client.Headers[HttpRequestHeader.Authorization] = "Bearer " + APIKEY;
            //client.BaseAddress = url;

            //Console.WriteLine("Fetching {0}", url);
            try
            {
                client.DownloadFile(url, filePath);
                return true;
            }
            catch (WebException e)
            {
                String msg = e.Message.ToString() + " ";
                if (e.Response != null)
                {
                    using (WebResponse response = e.Response)
                    {
                        Stream stream = response.GetResponseStream();
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            msg += reader.ReadToEnd();
                        }
                    }
                }
                Console.WriteLine("Error: {0}", msg);
                return false;
            }
        }

        static string Fetch(WebClient client, string url)
        {
            //Console.WriteLine("Fetching {0}", url);
            try
            {
                return client.DownloadString(url);
            }
            catch (WebException e)
            {
                String msg = e.Message.ToString() + " ";
                if (e.Response != null)
                {
                    using (WebResponse response = e.Response)
                    {
                        Stream stream = response.GetResponseStream();
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            msg += reader.ReadToEnd();
                        }
                    }
                }
                Console.WriteLine("Error: {0}", msg);
                return null;
            }
        }

        static string FetchTMDB(string apikey, string query)
        {
            string url = String.Format("https://api.themoviedb.org/3{0}", query);

            WebClient client = new WebClient();
            client.Headers[HttpRequestHeader.ContentType] = "application/json";
            client.Headers[HttpRequestHeader.Authorization] = "Bearer " + apikey;
            //client.BaseAddress = url;

            return Fetch(client, url);
        }

        static string[] FetchPosters(string cover_path, string wmcid, string poster_path)
        {
            string extension = Path.GetExtension(poster_path);

            string small = String.Format("https://image.tmdb.org/t/p/{0}{1}", IMGWIDTH_SMALL, poster_path);
            string large = String.Format("https://image.tmdb.org/t/p/{0}{1}", IMGWIDTH_LARGE, poster_path);
            string filename_s = String.Format("{0}_small{1}", wmcid, extension);
            string filename_l = String.Format("{0}_large{1}", wmcid, extension);

            string filepath_s = Path.Combine(cover_path, filename_s);
            string filepath_l = Path.Combine(cover_path, filename_l);

            string[] result = new string[2];
            // skip over already downloaded images
            result[0] = (File.Exists(filepath_s) ? filename_s : (FetchData(small, filepath_s) ? filename_s : null));
            result[1] = (File.Exists(filepath_l) ? filename_l : (FetchData(small, filepath_l) ? filename_l : null));
            // cleanup
            if (result[1] == null && result[0] != null)
                result[1] = result[0];
            if (result[0] == null && result[1] != null)
                result[0] = result[1];
            return result;
        }

        static string[][] entities = new string[][]
        {
            new string[] { "%", "%25" },
            new string[] { "+", "%2B" },
            new string[] { "&", "%26" },
            new string[] { "#", "%23" }
        };

        static Movie FetchMovie(string apikey, string title, string year)
        {
            // encoded away entities for url query parameter
            string encoded = title;
            foreach (string[] entity in entities)
                encoded = encoded.Replace(entity[0], entity[1]);

            // try primary_release_year first
            string query = String.Format("/search/movie?query={0}&include_adult=true&language=en-US&page=1&primary_release_year={1}", encoded, year);
            string lookup = FetchTMDB(apikey, query);
            if (String.IsNullOrEmpty(lookup))
                return null;

            Page page = FromJSON<Page>(lookup);
            if (page.total_results == 0)
            {
                // try year second
                query = String.Format("/search/movie?query={0}&include_adult=true&language=en-US&page=1&year={1}", encoded, year);
                lookup = FetchTMDB(apikey, query);
                if (String.IsNullOrEmpty(lookup))
                    return null;

                page = FromJSON<Page>(lookup);
                if (page.total_results == 0)
                    return null;
            }

            Result result = page.results[0];

            string info = String.Format("/movie/{0}?language=en-US&append_to_response=release_dates,credits", result.id);
            string json = FetchTMDB(apikey, info);
            if (String.IsNullOrEmpty(json))
                return null;

            return FromJSON<Movie>(json);
        }


        static void WriteXML<T>(T obj, string path)
        {
            XmlSerializer mySerializer = new XmlSerializer(typeof(T));
            StreamWriter myWriter = new StreamWriter(path);
            mySerializer.Serialize(myWriter, obj);
            myWriter.Close(); 
        }


        static void Usage()
        {
            Console.WriteLine("Usage: {0} srcMovieFolder dstLibraryFolder [/server|API_KEY|apikey.txt] [/startFromNum|individualMovieNames ...]", Environment.CommandLine);
            Environment.Exit (1);
        }

        static string GetYear(string title)
        {
            int end = title.LastIndexOf(')');
            if (end == -1)
                return null;
            int start = title.LastIndexOf('(', end - 1);
            if (start == -1)
                return null;
            return title.Substring(start + 1, end - start - 1);
        }

        //static string FixFilename(string filename)
        //{
        //    StringBuilder buf = new StringBuilder();
        //    foreach (char c in filename)
        //    {
        //        if ((int)c > 0xFF)
        //            buf.Append('~');
        //        else
        //            buf.Append(c);
        //    }
        //    return buf.ToString();
        //}

        static string FixTitle(string filename)
        {
            StringBuilder buf = new StringBuilder();
            foreach (char c in filename)
            {
                switch (c)
                {
                    case 'ː':
                        buf.Append(':');
                        break;
                    case '~':
                        buf.Append('/');
                        break;
                    default:
                        buf.Append(c);
                        break;
                }
            }
            return buf.ToString();
        }

        //static Dictionary<string, List<string>> ratings = new Dictionary<string, List<string>>();

        //static void AddRating(string country, string rating)
        //{
        //    rating = rating.Trim();
        //    if (String.IsNullOrEmpty(rating))
        //        return;

        //    rating = String.Format("'{0}'", rating);

        //    List<string> list = null;
        //    if (!ratings.ContainsKey(country))
        //        ratings.Add(country, (list = new List<string>()));
        //    else
        //        list = ratings[country];

        //    if (!list.Contains(rating))
        //        list.Add(rating);
        //}

        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            ServicePointManager.SecurityProtocol = Tls12;

            if (args.Length < 3)
            {
                Usage();
                return;
            }

            string src = args[0];
            string dst = args[1];
            bool server = (args[2].Equals("/server", StringComparison.CurrentCultureIgnoreCase));
            string apikey = (server ? "" : args[2]);

            if (!server && File.Exists(apikey))
            {
                // read apikey from file
                string contents = File.ReadAllText(apikey);
                if (String.IsNullOrEmpty(contents))
                    throw new Exception("ApiKey text file '" + apikey + "' is empty!");

                apikey = contents;
            }

            if (!Directory.Exists(src))
                throw new Exception("Source folder '" + src + "' is invalid!");

            if (!Directory.Exists(dst))
                throw new Exception("Destination folder '" + dst + "' is invalid!");


            string data = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string microsoft = Path.Combine(data, "Microsoft");
            string ehome = Path.Combine(microsoft, "eHome");

            if (!server && !Directory.Exists(ehome))
                throw new Exception("eHome folder '" + ehome + "' is invalid!");

            string covers = Path.Combine(ehome, "DvdCoverCache");
            string wmcids = Path.Combine(ehome, "DvdInfoCache");

            List<string> missing = new List<string>();

            string extension = ".mp4";
            string filter = "* (*)" + extension;
            string[] files = Directory.GetFiles(src, filter, SearchOption.AllDirectories);
            int skip = 0;
            int index = 0;
            int known = 0;
            bool specific = (args.Length > 3);
            if (specific && args.Length == 4 && args[3].StartsWith("/"))
            {
                string token = args[3].Substring(1).Trim();
                if (!int.TryParse(token, out skip))
                {
                    Usage();
                    return;
                }
                // disable specific
                specific = false;
            }
            int count = (specific ? args.Length - 3 : files.Length);
            foreach (string file in files)
            {
                string folder = Path.GetDirectoryName(file);

                // chop off suffix's and collect year
                string filename = Path.GetFileName(file);
                string name = filename.Substring(0, filename.Length - extension.Length).Trim();
                string title = FixTitle(name);
                string year = GetYear(title);
                string next = year;
                while (!String.IsNullOrEmpty(next))
                {
                    // +2 for brackets
                    title = title.Substring(0, title.Length - (next.Length + 2)).Trim();
                    next = GetYear(title);
                    if (!String.IsNullOrEmpty(next))
                        year = next;
                }

                string path = Path.Combine(dst, name);//FixFilename(name));
                string dest = Path.Combine(path, filename);//FixFilename(filename));

                string srt = String.Format("{0}.srt", name);
                string subtitles = Path.Combine(folder, srt);
                string clone = Path.Combine(path, srt);

                if (specific)
                {
                    // check for 'name' match
                    bool match = false;
                    for (int i = 3; i < args.Length; i++)
                        if (args[i].Equals(name, StringComparison.CurrentCultureIgnoreCase))
                        {
                            match = true;
                            break;
                        }
                    if (!match)
                        continue;
                }

                if (skip > ++index)
                    continue;

                Console.WriteLine("[{0}/{1}] '{2}' -> '{3}'", index, count, file, dest);

                if (server)
                {
                    if (!Directory.Exists(path))
                    {
                        Console.WriteLine("Movie metadata has not been updated for '{0}'", title);
                        missing.Add(title);
                        continue;
                    }

                    // server mode creates hard links to original movies
                    if (!File.Exists(dest))
                    {
                        bool success = CreateHardLink(dest, file, IntPtr.Zero);
                        if (!success)
                            throw new Exception("Could not create movie hardlink '" + dest + "'");
                    }

                    if (File.Exists(subtitles) && !File.Exists(clone))
                    {
                        bool success = CreateHardLink(clone, subtitles, IntPtr.Zero);
                        if (!success)
                            throw new Exception("Could not create subtitle hardlink '" + clone + "'");
                    }
                }
                else
                {
                    // grab json from TMBD
                    Movie movie = FetchMovie(apikey, title, year);
                    if (movie == null)
                    {
                        Console.WriteLine("Unable to fetch movie '{0}'", title);
                        missing.Add(title);
                        continue;
                    }

                    //// collect ratings for development
                    //foreach (Releases releases in movie.release_dates.results)
                    //    foreach (Release release in releases.release_dates)
                    //        AddRating(releases.iso_3166_1, release.certification);
                    
                    if (!Directory.Exists(path))
                    {
                        DirectoryInfo dir = Directory.CreateDirectory(path);
                        if (!dir.Exists)
                            throw new Exception("Could not create movie folder '" + path + "'");
                    }

                    // ensure eHome folders exist
                    Directory.CreateDirectory(covers);
                    Directory.CreateDirectory(wmcids);

                    string wmcid = String.Format("wmc_id__{0}", movie.id);

                    string[] imgnames = FetchPosters(covers, wmcid, movie.poster_path);
                    if (String.IsNullOrEmpty(imgnames[0]) || String.IsNullOrEmpty(imgnames[1]))
                        Console.WriteLine("Unable to update poster for '{0}'", movie.title);

                    // create xml files....
                    string dvdxml = Path.Combine(path, String.Format("{0}.dvdid.xml", wmcid));
                    WriteXML<DVDID>(DVDID.Create(wmcid, title), dvdxml);

                    string wmcxml = Path.Combine(wmcids, String.Format("{0}.xml", wmcid));
                    WriteXML<WMCID>(WMCID.Create(wmcid, MDRDVD.Create(movie, imgnames)), wmcxml);
                }

                known = index;
            }

            foreach (string missed in missing)
                Console.WriteLine("Missed {0}", missed);

            if (!specific && known != index)
                Console.WriteLine("Last known {0}", known);

            //// for development
            //Console.WriteLine("Ratings:");
            //foreach (string country in ratings.Keys)
            //{
            //    Console.WriteLine("  Country: '{0}'", country);
            //    Console.Write("    Ratings: ");
            //    foreach (string rating in ratings[country])
            //        Console.Write("{0}; ", rating);
            //    Console.WriteLine();
            //}

            Console.WriteLine("Done");
        }
    }
}
