//本文件为自动生成，请勿手动修改
using System.Collections.Generic;
public partial class test2
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
		t1 = 2,
	};

	public static List<TestT> tt1 = new List<TestT>()
	{
		new TestT
		{
			t0 = 1,
			t1 = 2,
		},
		new TestT
		{
			t0 = 3,
			t1 = 4,
		},
	};

}
public class TestT
{
	public int t0;
	public int t1;
}


