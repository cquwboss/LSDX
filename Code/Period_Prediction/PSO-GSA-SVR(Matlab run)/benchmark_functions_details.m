%PSOGSA source code v3.0, Generated by SeyedAli Mirjalili, 2011. 
%Adopted from: S. Mirjalili, S.Z. Mohd Hashim, �A New Hybrid PSOGSA 
%Algorithm for Function Optimization, in IEEE International Conference 
%on Computer and Information Application?ICCIA 2010), China, 2010, pp.374-377.

% This function gives boundaries and dimension of search space for test functions.
function [down,up,dim]=benchmark_functions_details(Benchmark_Function_ID)

%If lower bounds of dimensions are the same, then 'down' is a value.
%Otherwise, 'down' is a vector that shows the lower bound of each dimension.
%This is also true for upper bounds of dimensions.

%Insert your own boundaries with a new Benchmark_Function_ID.

dim=30;
if Benchmark_Function_ID==1
    down=-100;up=100;
end

if Benchmark_Function_ID==2
    down=-10;up=10;
end

if Benchmark_Function_ID==3
    down=-100;up=100;
end

if Benchmark_Function_ID==4
    down=-100;up=100;
end

if Benchmark_Function_ID==5
    down=-30;up=30;
end

if Benchmark_Function_ID==6
    down=-100;up=100;
end

if Benchmark_Function_ID==7
    down=-1.28;up=1.28;
end

if Benchmark_Function_ID==8
    down=-500;up=500;
end

if Benchmark_Function_ID==9
    down=-5.12;up=5.12;
end

if Benchmark_Function_ID==10
    down=-32;up=32;
end

if Benchmark_Function_ID==11
    down=-600;up=600;
end

if Benchmark_Function_ID==12
    down=-50;up=50;
end

if Benchmark_Function_ID==13
    down=-50;up=50;
end

if Benchmark_Function_ID==14
    down=-65.536;up=65.536;dim=2;
end

if Benchmark_Function_ID==15
    down=-5;up=5;dim=4;
end

if Benchmark_Function_ID==16
    down=-5;up=5;dim=2;
end

if Benchmark_Function_ID==17
    down=[-5;0];up=[10;15];dim=2;
end

if Benchmark_Function_ID==18
    down=-2;up=2;dim=2;
end

if Benchmark_Function_ID==19
    down=0;up=1;dim=3;
end

if Benchmark_Function_ID==20
    down=0;up=1;dim=6;
end

if Benchmark_Function_ID==21
    down=0;up=10;dim=4;
end

if Benchmark_Function_ID==22
    down=0;up=10;dim=4;
end

if Benchmark_Function_ID==23
    down=0;up=10;dim=4;
end
%--SVR 
if Benchmark_Function_ID==24
    down=1;up=100;dim=2;
end