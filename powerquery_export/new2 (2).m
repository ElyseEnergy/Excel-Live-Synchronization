let
    Source = Excel.Workbook(Web.Contents("https://eu3.ragic.com/elyseenergy/projets/2.xlsx?APIKey=Njl3OENtYnFnTExxSzNWVXZ6Y2E1Tlg0RWtjcVVBdnE2SkxsM2pVWnVJNnFEb0d1SDV6cVRJNytaU09CaHR3MjlzVlVFK1lvR09NPQ=="), null, true),
    Projets_Sheet = Source{[Item="Projets",Kind="Sheet"]}[Data]
in
    Projets_Sheet