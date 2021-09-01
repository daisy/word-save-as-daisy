using System.Collections;

namespace Daisy.SaveAsDAISY.Conversion
{
	/// <summary>
	///  Chained resource managers
	/// </summary>
	public class ChainResourceManager : System.Resources.ResourceManager
	{
		private readonly ArrayList managers;

		public ChainResourceManager()
		{
			this.managers = new ArrayList();
		}

		public void Add(System.Resources.ResourceManager manager)
		{
			managers.Add(manager);
		}

		/// <summary>
		/// Function which return value of a particular key
		/// </summary>
		/// <param name="key"></param>
		/// <returns></returns>
		public override string GetString(string key)
		{
			for (int i = this.managers.Count - 1; i >= 0; i--)
			{
				System.Resources.ResourceManager manager = (System.Resources.ResourceManager)this.managers[i];
				if (manager != null)
				{
					string value = manager.GetString(key);
					if (value != null && value.Length > 0)
					{
						return value;
					}
				}
			}
			return null;
		}
	}
}