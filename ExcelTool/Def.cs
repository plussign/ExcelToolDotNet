public class Def
{
    public static string LINE = "/bbb";
    public static string CELL = "/ccc";

    public const string ERL_IMPL_BEGIN = @"-module(csv).


-include(""csv.hrl"").

-export([cfg_file_list/0]).
-export([load_csv_to_ets/1]).
-export([load_csv/2]).



-define(RES_LINE, ""/bbb"").
-define(RES_CELL, ""/ccc"").
load_csv_to_ets(Path) ->
	lists:foreach(
		fun(FileName) ->
			load_csv(FileName, Path)
		end, cfg_file_list()).

to_list(V) -> V.

to_float(V) when is_list(V) -> 
	case lists:member($., V) of
		true -> list_to_float(V);
		false -> list_to_integer(V)
	end;
to_float(V) -> V.

to_int(V) when is_list(V) -> list_to_integer(V);
to_int(V) -> V.

read_csv_line(<<"""">>) -> [];read_csv_line(FileData) ->
	lists:map(
		fun(Line) ->
			[binary_to_list(Item) || Item <- re:split(Line, ?RES_CELL)]
		end, re:split(FileData, ?RES_LINE)).
cfg_file_list() ->
	L = [
";


    public const string ERL_IMPL_END = @"""""
	],
	[_|T] = lists:reverse(L),
	lists:reverse(T).
";

}
