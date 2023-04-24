
set /p input= Type any input:
FOR /L  %%a IN (w,1,%input%) DO   start cmd /k  java -jar "C:\Users\Alexandru Andercou\Desktop\CPD_AlgDistr\BullyAlgorithm\out\production\BullyAlgorithm\BullyAlg.jar"  %%a
