//本文件为自动生成，请勿手动修改
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

	public static TestT OnGetTestTFrom_tt1(int t0)
	{
		System.Diagnostics.Debug.Assert(tt1.ContainsKey(t0), $"Invalid t0 {t0}");
		return tt1[t0];
	}
	public static Dictionary<int, TestT> tt1 = new Dictionary<int, TestT>()
	{
		{
			1
			, new TestT
			{
				t0 = 1,
				t1 = "testt1",
				t2 = new int[]{1,2,3,4,5,6},
				t3 = new string[]{"t1","t2","t4"},
				t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d2"}},
				t5 = false,
				t6 = 12.1f,
				t7 = 12.222,
			}
		},
		{
			2
			, new TestT
			{
				t0 = 2,
				t1 = "testt2",
				t2 = new int[]{1,2,3,4,5,7},
				t3 = new string[]{"t1","t2","t5"},
				t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
				t5 = false,
				t6 = 12.2f,
				t7 = 13.222,
			}
		},
		{
			12
			, new TestT
			{
				t0 = 12,
				t1 = "testt12",
				t2 = new int[]{1,2,3,4,5,7},
				t3 = new string[]{"t1","t2","t5"},
				t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
				t5 = true,
				t6 = 120.2f,
				t7 = 130.222,
			}
		},
	};

	public static TestT OnGetTestTFrom_dTestT(int t0)
	{
		System.Diagnostics.Debug.Assert(dTestT.ContainsKey(t0), $"Invalid t0 {t0}");
		return dTestT[t0];
	}
	public static Dictionary<int, TestT> dTestT = new Dictionary<int, TestT>()
	{
		{
			1
			, new TestT
			{
				t0 = 1,
				t1 = "testt1",
				t2 = new int[]{1,2,3,4,5,6},
				t3 = new string[]{"t1","t2","t4"},
				t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d2"}},
				t5 = false,
				t6 = 12.1f,
				t7 = 12.222,
			}
		},
		{
			2
			, new TestT
			{
				t0 = 2,
				t1 = "testt2",
				t2 = new int[]{1,2,3,4,5,7},
				t3 = new string[]{"t1","t2","t5"},
				t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
				t5 = false,
				t6 = 12.2f,
				t7 = 13.222,
			}
		},
		{
			12
			, new TestT
			{
				t0 = 12,
				t1 = "testt12",
				t2 = new int[]{1,2,3,4,5,7},
				t3 = new string[]{"t1","t2","t5"},
				t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
				t5 = true,
				t6 = 120.2f,
				t7 = 130.222,
			}
		},
		{
			13
			, new TestT
			{
				t0 = 13,
				t1 = "testt12",
				t2 = new int[]{1,2,3,4,5,7},
				t3 = new string[]{"t1","t2","t5"},
				t4 = new Dictionary<int, string>(){{1,"d1"}, {2, "d3"}},
				t5 = true,
				t6 = 120.2f,
				t7 = 130.222,
			}
		},
	};

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
	};

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


