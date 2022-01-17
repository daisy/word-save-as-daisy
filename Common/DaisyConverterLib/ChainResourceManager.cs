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
		/// <param name="name"></param>
		/// <returns></returns>
		public override string GetString(string name)
		{
			for (int i = this.managers.Count - 1; i >= 0; i--)
			{
				System.Resources.ResourceManager manager = (System.Resources.ResourceManager)this.managers[i];
				if (manager != null)
				{
                    try {
						string value = manager.GetString(name);
						if (value != null && value.Length > 0) {
							return value;
						}
					} catch (System.Exception) {
						// value not found in this manager
                    }
					
				}
			}
			// string not found, avoid raising hidden exception by returning the key
			return name;
		}
	}
}