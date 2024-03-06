using Newtonsoft.Json;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class JsonTest : MonoBehaviour
{
    // Start is called before the first frame update
    void Start()
    {
        Debug.Log(item.Get(2).res);
        Debug.Log(Skill.Get(1001).Skill_S2);
        Vector2 vector2 = new Vector2(Skill.Get(1001).Skill_S2[0], Skill.Get(1001).Skill_S2[1]

            );
        Debug.Log(vector2);
    }


}
