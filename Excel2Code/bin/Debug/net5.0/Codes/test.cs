//本文件为自动生成，请勿手动修改
//--------------------------
//https://github.com/yukuyoulei/Excel2CSharp
//--------------------------
//
using System.Collections.Generic;
public partial class test
{
	public static int test1 = 1;
	public static int test2 = 2;
	public static int test3 = 3;
	public static int test4 = 4;
	public static string test5 = "testdesc";
	public static string test6 = $"simple {test5}";
	public static float Test7 = 1.2f;
	public static bool Test8 = true;

	public static TestT tt0 = new TestT()
	{
		t0 = 1,
		t1 = "testt1",
		t2 = new int[]{1,2,3,4,5,6},
		t3 = new string[]{"t1","t2","t4"},
		t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d2"}},
		t5 = false,
		t6 = 12.1f,
		t7 = 12.222,
	};

	public static Dictionary<int, TestT> tt1 = new Dictionary<int, TestT>();
	public static TestT OnGetFrom_tt1(int t0)
	{
		switch (t0)
		{
			case 1:
				if (!tt1.ContainsKey(1))
				{
					var data = new TestT()
					{
						t0 = 1,
						t1 = "testt1",
						t2 = new int[]{1,2,3,4,5,6},
						t3 = new string[]{"t1","t2","t4"},
						t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d2"}},
						t5 = false,
						t6 = 12.1f,
						t7 = 12.222,
					};
					tt1[1] = data;
				}
				return tt1[1];
			case 2:
				if (!tt1.ContainsKey(2))
				{
					var data = new TestT()
					{
						t0 = 2,
						t1 = "testt2",
						t2 = new int[]{1,2,3,4,5,7},
						t3 = new string[]{"t1","t2","t5"},
						t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
						t5 = false,
						t6 = 12.2f,
						t7 = 13.222,
					};
					tt1[2] = data;
				}
				return tt1[2];
			case 12:
				if (!tt1.ContainsKey(12))
				{
					var data = new TestT()
					{
						t0 = 12,
						t1 = "testt12",
						t2 = new int[]{1,2,3,4,5,7},
						t3 = new string[]{"t1","t2","t5"},
						t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
						t6 = 120.2f,
						t7 = 130.222,
					};
					tt1[12] = data;
				}
				return tt1[12];
		}
		return null;
	}
	public static List<int> allTestTs = new List<int>(){1,2,12,};

	public static Dictionary<int, TestT> dTestT = new Dictionary<int, TestT>();
	public static TestT OnGetFrom_dTestT(int t0)
	{
		switch (t0)
		{
			case 1:
				if (!dTestT.ContainsKey(1))
				{
					var data = new TestT()
					{
						t0 = 1,
						t1 = "testt1",
						t2 = new int[]{1,2,3,4,5,6},
						t3 = new string[]{"t1","t2","t4"},
						t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d2"}},
						t5 = false,
						t6 = 12.1f,
						t7 = 12.222,
					};
					dTestT[1] = data;
				}
				return dTestT[1];
			case 2:
				if (!dTestT.ContainsKey(2))
				{
					var data = new TestT()
					{
						t0 = 2,
						t1 = "testt2",
						t2 = new int[]{1,2,3,4,5,7},
						t3 = new string[]{"t1","t2","t5"},
						t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
						t5 = false,
						t6 = 12.2f,
						t7 = 13.222,
					};
					dTestT[2] = data;
				}
				return dTestT[2];
			case 12:
				if (!dTestT.ContainsKey(12))
				{
					var data = new TestT()
					{
						t0 = 12,
						t1 = "testt12",
						t2 = new int[]{1,2,3,4,5,7},
						t3 = new string[]{"t1","t2","t5"},
						t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
						t5 = true,
						t6 = 120.2f,
						t7 = 130.222,
					};
					dTestT[12] = data;
				}
				return dTestT[12];
			case 13:
				if (!dTestT.ContainsKey(13))
				{
					var data = new TestT()
					{
						t0 = 13,
						t1 = "testt12",
						t2 = new int[]{1,2,3,4,5,7},
						t3 = new string[]{"t1","t2","t5"},
						t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
						t5 = true,
						t6 = 120.2f,
						t7 = 130.222,
					};
					dTestT[13] = data;
				}
				return dTestT[13];
		}
		return null;
	}
	public static List<int> allTestTs = new List<int>(){1,2,12,13,};

	public static List<TestT> lTests = new List<TestT>()
	{
		new TestT
		{
			t0 = 1,
			t1 = "testt1",
			t2 = new int[]{1,2,3,4,5,6},
			t3 = new string[]{"t1","t2","t4"},
			t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d2"}},
			t5 = false,
			t6 = 12.1f,
			t7 = 12.222,
		},
		new TestT
		{
			t0 = 2,
			t1 = "testt2",
			t2 = new int[]{1,2,3,4,5,7},
			t3 = new string[]{"t1","t2","t5"},
			t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
			t5 = false,
			t6 = 12.2f,
			t7 = 13.222,
		},
		new TestT
		{
			t0 = 12,
			t1 = "testt12",
			t2 = new int[]{1,2,3,4,5,7},
			t3 = new string[]{"t1","t2","t5"},
			t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
			t5 = true,
			t6 = 120.2f,
			t7 = 130.222,
		},
		new TestT
		{
			t0 = 13,
			t1 = "testt12",
			t2 = new int[]{1,2,3,4,5,7},
			t3 = new string[]{"t1","t2","t5"},
			t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
			t5 = true,
			t6 = 120.2f,
			t7 = 130.222,
		},
	}

}
public class TestT
{
	public int t0; //t0
	public string t1; //t1
	public int[] t2; //t2
	public string[] t3; //t3
	public Dictionary<int, string> t4; //t4
	public bool t5; //t5
	public float t6; //t6
	public double t7; //t7
}


