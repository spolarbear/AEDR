/*
 * Created by SharpDevelop.
 * User: Yangxin
 * Date: 2015/1/18
 * Time: 12:44
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

using System;
using System.IO;
using System.Text;

namespace sunproject
{
	/// <summary>
	/// Description of midasResultFetch.
	/// </summary>
	public class midasResultFetch
	{
		private string path;
		public midasResultFetch(string path)
		{
			this.path=path;
		}
		public void setPath(string nPath){
			this.path=nPath;
		}
		private void fileFetch(string path){
		    StreamReader sr = new StreamReader(path,Encoding.Default);
		    String line;
		    while ((line = sr.ReadLine()) != null) 
		    {
		    	Console.WriteLine(line.ToString());
		    }
		}
	}

}
