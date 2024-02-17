#See https://aka.ms/customizecontainer to learn how to customize your debug container and how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/aspnet:6.0 AS base
RUN apt-get update && apt-get install -y --allow-unauthenticated libgdiplus
WORKDIR /app
EXPOSE 80


FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src

ENV ASPNETCORE_ENVIRONMENT Production

COPY ["BS_FileGenerator/BS_FileGenerator.csproj", "BS_FileGenerator/"]
RUN dotnet restore "BS_FileGenerator/BS_FileGenerator.csproj"
COPY . .
WORKDIR "/src/BS_FileGenerator"
RUN dotnet build "BS_FileGenerator.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "BS_FileGenerator.csproj" -c Release -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "BS_FileGenerator.dll"]