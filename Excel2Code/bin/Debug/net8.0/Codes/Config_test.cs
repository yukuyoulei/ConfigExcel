//本文件为自动生成，请勿手动修改
//
using System;
using System.Collections.Generic;
public partial class Config_test
{
	public const int test1 = 1;
	public const int test2 = 2;
	public const int test3 = 3;
	public const int test4 = 4;
	public const string test5 = "testdesc";
	public const float Test7 = 1.2f;
	public static bool Test8 = true;

	public static Dictionary<int, TestT> dTestT = new Dictionary<int, TestT>();
	public static TestT OnGetFrom_dTestT(int id)
	{
		TestT data = null;
		if (dTestT.TryGetValue(id, out data))
		{
			return data;
		}
		var t = typeof(Config_test);
		var m = t.GetMethod($"CreateTestT_{id}", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
		if (m == null) return null;
		data = m.Invoke(null, null) as TestT;
		dTestT[id] = data;
		return data;
	}
	private static TestT CreateTestT_1()
	{
		return new TestT()
		{
			t0 = 1,
			t1 = @"testt1",
			t2 = new int[]{1,2,3,4,5,6},
			t3 = new string[]{"t1","t2","t4"},
			t4 = new Dictionary<int,string>(){{1,"d1"}, {2, "d2"}},
			t5 = false,
			t6 = 12.1f,
			t7 = 12.222,
		};
	}
	private static TestT CreateTestT_2()
	{
		return new TestT()
		{
			t0 = 2,
			t1 = @"testt2",
			t2 = new int[]{1,2,3,4,5,7},
			t3 = new string[]{"t1","t2","t5"},
			t4 = new Dictionary<int,string>(){{1,"d1"}, {2, "d3"}},
			t5 = false,
			t6 = 12.2f,
			t7 = 13.222,
		};
	}
	private static TestT CreateTestT_12()
	{
		return new TestT()
		{
			t0 = 12,
			t1 = @"testt12",
			t2 = new int[]{1,2,3,4,5,7},
			t3 = new string[]{"t1","t2","t5"},
			t4 = new Dictionary<int,string>(){{1,"d1"}, {2, "d3"}},
			t5 = true,
			t6 = 120.2f,
			t7 = 130.222,
		};
	}
	private static TestT CreateTestT_13()
	{
		return new TestT()
		{
			t0 = 13,
			t1 = @"testt12",
			t2 = new int[]{1,2,3,4,5,7},
			t3 = new string[]{"t1","t2","t5"},
			t4 = new Dictionary<int,string>(){{1,"d1"}, {2, "d3"}},
			t5 = true,
			t6 = 120.2f,
			t7 = 130.222,
		};
	}
	public static List<int> allTestTs = new List<int>()
		{
			1,
			2,
			12,
			13,

		};
	public static TestT OnGetFrom_TestTByIndex(int index){if (index < 0 || index >= allTestTs.Count) return null; return OnGetFrom_dTestT(allTestTs[index]);}

}
public partial class TestT
{
	public int t0; /*t0*/
	public string t1; /*t1*/
	public int[] t2; /*t2*/
	public string[] t3; /*t3*/
	public Dictionary<int,string> t4; /*t4*/
	public bool t5; /*t5*/
	public float t6; /*t6*/
	public double t7; /*t7*/
}


