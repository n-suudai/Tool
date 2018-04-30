using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDiffTool
{

    public class Diff
    {
        const long KB = 1024;
        const long MB = KB * 1024;
        const long GB = MB * 1024;

        public enum Status
        {
            None,
            Add,
            Delete,
        }

        public class Line
        {
            public int    Index = -1;
            public string Text = "";
            public Status Status = Status.None;

            public string StatusText
            {
                get
                {
                    string text = "|";
                    if (Status == Status.Add)
                    {
                        text = "+";
                    }
                    else if(Status == Status.Delete)
                    {
                        text = "-";
                    }

                    return Index + " " + text + " " + Text;
                }
            }
        }
        
        private class Tree
        {
            public Tree(Status status, Tree prev)
            {
                Status = status;
                Prev = prev;
            }

            public void Set(Status status, Tree prev)
            {
                Status = status;
                Prev = prev;
            }

            public Tree Prev = null;
            public Status Status = Status.None;
        }

        private class FP
        {
            public int Y = -1;
            public Tree Tree = null;
        }

        public static List<Line> Execute(List<string> a, List<string> b)
        {
            List<string> A = a;
            List<string> B = b;
            int M = a.Count;
            int N = b.Count;
            bool Exchanged = false;

            if(M == 0 && N == 0)
            {
                return new List<Line>();
            }

            // 長さによって入れ替え
            if (N > M)
            {
                Exchanged = true;
                A = b;
                B = a;

                int Temp = M;
                M = N;
                N = Temp;
            }

            int offset = N;
            int delta = M - N;
            int size = M + N + 1;

            FP[] fp = new FP[size];
            for (int i = 0; i < size; ++i)
            {
                fp[i] = new FP();
            }

            Action<int> snake = (int k) =>
            {
                var current = fp[k + offset];

                if (k == -N)
                {
                    // 横が存在しない
                    // down  => 上から降りてくる => ADD
                    // down の y座標は上から降りてくるので+1
                    var down = fp[k + 1 + offset];
                    current.Y = down.Y + 1;

                    current.Tree = new Tree(Status.Add, down.Tree);
                }
                else if (k == M)
                {
                    // 上が存在しない
                    // slide => 横からスライド   => DEL
                    var slide = fp[k - 1 + offset];
                    current.Y = slide.Y;

                    current.Tree = new Tree(Status.Delete, slide.Tree);
                }
                else
                {
                    var slide = fp[k - 1 + offset];
                    var down = fp[k + 1 + offset];
                    if (slide.Y == -1 && down.Y == -1)
                    {
                        // どちらも未定義 => (0, 0)について
                        current.Y = 0;
                    }
                    else if (down.Y == -1 || slide.Y == -1)
                    {
                        // どちらかが未定義状態
                        if (down.Y == -1)
                        {
                            current.Y = slide.Y;

                            current.Tree = new Tree(Status.Delete, slide.Tree);
                        }
                        else
                        {
                            current.Y = down.Y + 1;

                            current.Tree = new Tree(Status.Add, down.Tree);
                        }
                    }
                    else
                    {
                        // どちらも定義済み
                        if (slide.Y > (down.Y + 1))
                        {
                            current.Y = slide.Y;

                            current.Tree = new Tree(Status.Delete, slide.Tree);
                        }
                        else
                        {
                            current.Y = down.Y + 1;

                            current.Tree = new Tree(Status.Add, down.Tree);
                        }
                    }
                }
                int y = current.Y;
                int x = y + k;
                while (x < M && y < N && A[x] == B[y])
                {
                    current.Tree = new Tree(Status.None, current.Tree);
                    x++;
                    y++;
                }
                current.Y = y;
            };

            {
                int p = -1, k = 0;
                // k = deltaのときのy座標が fp[delta+offset]
                // 目的地はk = delta上にあるので,
                // fp[delta+offset]のy座標がNであるとき, x座標もMであり, 目的地
                while (fp[delta + offset].Y < N)
                {
                    p = p + 1;
                    for (k = -p; k < delta; ++k)
                    {
                        snake(k);
                    }
                    for (k = delta + p; k > delta; --k)
                    {
                        snake(k);
                    }

                    snake(delta);
                }
            }

            List<Line> list = new List<Line>();
            {
                var current = fp[delta + offset];
                //current.d = delta + 2 * p;
                // tree 完成後処理
                // treeの終端から先頭に向かってたどって行く
                // prevをたどって後ろから戻して行く
                
                Status status = Status.None;

                int a_index = M - 1, b_index = N - 1;
                for (var i = current.Tree; i != null; i = i.Prev)
                {
                    // "+" or   "-"  or "|"
                    // add or delete or common
                    status = i.Status;

                    Line line = new Line();
                    list.Insert(0, line);

                    // exchangedも考慮するように
                    if (status == Status.Add)
                    {
                        // addのとき, Bを調べて足された文字を調べる
                        // ひっくり返されてたら, delにする
                        line.Index = b_index;
                        line.Text = B[b_index];
                        line.Status = (Exchanged) ? Status.Delete : Status.Add;
                        --b_index;
                    }
                    else if (status == Status.Delete)
                    {
                        // delのとき, Aを調べて削られた文字を調べる
                        // ひっくり返されてたら, addにする
                        line.Index = a_index;
                        line.Text = A[a_index];
                        line.Status = (Exchanged) ? Status.Add : Status.Delete;
                        --a_index;
                    }
                    else
                    {
                        // commonのとき, どちらも同じなのでとりあえずAを調べている
                        line.Index = a_index;
                        line.Text = A[a_index];
                        line.Status = Status.None;
                        --a_index;
                        --b_index;
                    }
                }
            }
            return list;
        }

    }
}
