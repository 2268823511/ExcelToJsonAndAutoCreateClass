using System.Collections.Generic;
			{
				string jsonContent = File.ReadAllText(filePaths[0]);
				config = JsonConvert.DeserializeObject<Skill>(jsonContent);