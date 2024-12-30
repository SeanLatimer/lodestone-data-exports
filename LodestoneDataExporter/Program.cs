using FFXIV;
using Lumina;
using Lumina.Data;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using FlatSharp;
using Cyalume = Lumina.GameData;
using ISerializerExtensions = FlatSharp.ISerializerExtensions;

namespace LodestoneDataExporter
{
    public static class Program
    {
        private const string OutputDir = "../../../../pack";

        public static async Task Main(string[] args)
        {
            var dataPath = args.Length > 0
                ? args[0]
                : "C:/Program Files (x86)/SquareEnix/FINAL FANTASY XIV - A Realm Reborn/game/sqpack";
            var cyalume = new Cyalume(dataPath, new LuminaOptions { PanicOnSheetChecksumMismatch = false });

            await Task.WhenAll(
                Task.Run(() => ExportAchievementTable(cyalume)),
                Task.Run(() => ExportClassJobTable(cyalume)),
                Task.Run(() => ExportGuardianDeityTable(cyalume)),
                Task.Run(() => ExportGrandCompanyTable(cyalume)),
                Task.Run(() => ExportMinionTable(cyalume)),
                Task.Run(() => ExportMountTable(cyalume)),
                Task.Run(() => ExportRaceTable(cyalume)),
                Task.Run(() => ExportReputationTable(cyalume)),
                Task.Run(() => ExportTitleTable(cyalume)),
                Task.Run(() => ExportTownTable(cyalume)),
                Task.Run(() => ExportTribeTable(cyalume)),
                Task.Run(() => ExportItemTable(cyalume))
            );
        }

        private static void ExportAchievementTable(Cyalume cyalume)
        {
            var itemTable = new AchievementTable { Achievements = new List<Achievement>() };
            var languages = new[] { Language.English, Language.Japanese, Language.German, Language.French };
            foreach (var lang in languages)
            {
                var achievementSheet = cyalume.GetExcelSheet<Lumina.Excel.Sheets.Achievement>(lang);
                Parallel.ForEach(achievementSheet, new ParallelOptions { MaxDegreeOfParallelism = 4 }, achievement =>
                {
                    Achievement curAchievement;
                    lock (itemTable.Achievements)
                    {
                        curAchievement = itemTable.Achievements.FirstOrDefault(i => i.Id == achievement.RowId);
                        if (curAchievement == null)
                        {
                            curAchievement = new Achievement { Id = achievement.RowId };
                            itemTable.Achievements.Add(curAchievement);
                        }
                    }

                    switch (lang)
                    {
                        case Language.English:
                            curAchievement.NameEn = achievement.Name.ExtractText();
                            break;
                        case Language.Japanese:
                            curAchievement.NameJa = achievement.Name.ExtractText();
                            break;
                        case Language.German:
                            curAchievement.NameDe = achievement.Name.ExtractText();
                            break;
                        case Language.French:
                            curAchievement.NameFr = achievement.Name.ExtractText();
                            break;
                    }
                });
            }

            Serialize(
                Path.Join(OutputDir, "achievement_table.bin"),
                AchievementTable.Serializer,
                itemTable
            );
        }

        private static void ExportClassJobTable(Cyalume cyalume)
        {
            var classJobTable = new ClassJobTable { ClassJobs = new List<ClassJob>() };
            var languages = new[] { Language.English, Language.Japanese, Language.German, Language.French };
            foreach (var lang in languages)
            {
                var classJobSheet = cyalume.GetExcelSheet<Lumina.Excel.Sheets.ClassJob>(lang);
                foreach (var classJob in classJobSheet)
                {
                    var curClassJob = classJobTable.ClassJobs.FirstOrDefault(cj => cj.Id == classJob.RowId);
                    if (curClassJob == null)
                    {
                        curClassJob = new ClassJob
                            { Id = classJob.RowId, Parent = classJob.ClassJobParent.RowId, JobIndex = classJob.JobIndex };
                        classJobTable.ClassJobs.Add(curClassJob);
                    }

                    switch (lang)
                    {
                        case Language.English:
                            curClassJob.NameEn = classJob.Name.ExtractText();
                            break;
                        case Language.Japanese:
                            curClassJob.NameJa = classJob.Name.ExtractText();
                            break;
                        case Language.German:
                            curClassJob.NameDe = classJob.Name.ExtractText();
                            break;
                        case Language.French:
                            curClassJob.NameFr = classJob.Name.ExtractText();
                            break;
                    }
                }
            }

            Serialize(
                Path.Join(OutputDir, "classjob_table.bin"),
                ClassJobTable.Serializer,
                classJobTable
                );
        }

        private static void ExportGuardianDeityTable(Cyalume cyalume)
        {
            var deityTable = new DeityTable { Deities = new List<Deity>() };
            var languages = new[] { Language.English, Language.Japanese, Language.German, Language.French };
            foreach (var lang in languages)
            {
                var deitySheet = cyalume.GetExcelSheet<Lumina.Excel.Sheets.GuardianDeity>(lang);
                foreach (var deity in deitySheet)
                {
                    var curDeity = deityTable.Deities.FirstOrDefault(d => d.Id == deity.RowId);
                    if (curDeity == null)
                    {
                        curDeity = new Deity { Id = deity.RowId };
                        deityTable.Deities.Add(curDeity);
                    }

                    switch (lang)
                    {
                        case Language.English:
                            curDeity.NameEn = deity.Name.ExtractText();
                            break;
                        case Language.Japanese:
                            curDeity.NameJa = deity.Name.ExtractText();
                            break;
                        case Language.German:
                            curDeity.NameDe = deity.Name.ExtractText();
                            break;
                        case Language.French:
                            curDeity.NameFr = deity.Name.ExtractText();
                            break;
                    }
                }
            }

            Serialize(Path.Join(OutputDir, "deity_table.bin"), DeityTable.Serializer,deityTable);
        }

        private static void ExportGrandCompanyTable(Cyalume cyalume)
        {
            var gcTable = new GrandCompanyTable { GrandCompanies = new List<GrandCompany>() };
            var languages = new[] { Language.English, Language.Japanese, Language.German, Language.French };
            foreach (var lang in languages)
            {
                var gcSheet = cyalume.GetExcelSheet<Lumina.Excel.Sheets.GrandCompany>(lang);
                foreach (var gc in gcSheet)
                {
                    var curGc = gcTable.GrandCompanies.FirstOrDefault(c => c.Id == gc.RowId);
                    if (curGc == null)
                    {
                        curGc = new GrandCompany { Id = gc.RowId };
                        gcTable.GrandCompanies.Add(curGc);
                    }

                    switch (lang)
                    {
                        case Language.English:
                            curGc.NameEn = gc.Name.ExtractText();
                            break;
                        case Language.Japanese:
                            curGc.NameJa = gc.Name.ExtractText();
                            break;
                        case Language.German:
                            curGc.NameDe = gc.Name.ExtractText();
                            break;
                        case Language.French:
                            curGc.NameFr = gc.Name.ExtractText();
                            break;
                    }
                }
            }

            Serialize(Path.Join(OutputDir, "gc_table.bin"), GrandCompanyTable.Serializer, gcTable);
        }

        private static void ExportItemTable(Cyalume cyalume)
        {
            var itemTable = new ItemTable { Items = new List<Item>() };
            var languages = new[] { Language.English, Language.Japanese, Language.German, Language.French };
            foreach (var lang in languages)
            {
                var itemSheet = cyalume.GetExcelSheet<Lumina.Excel.Sheets.Item>(lang);
                Parallel.ForEach(itemSheet, new ParallelOptions { MaxDegreeOfParallelism = 4 }, item =>
                {
                    Item curItem;
                    lock (itemTable.Items)
                    {
                        curItem = itemTable.Items.FirstOrDefault(i => i.Id == item.RowId);
                        if (curItem == null)
                        {
                            curItem = new Item { Id = item.RowId };
                            itemTable.Items.Add(curItem);
                        }
                    }

                    switch (lang)
                    {
                        case Language.English:
                            curItem.NameEn = item.Name.ExtractText();
                            break;
                        case Language.Japanese:
                            curItem.NameJa = item.Name.ExtractText();
                            break;
                        case Language.German:
                            curItem.NameDe = item.Name.ExtractText();
                            break;
                        case Language.French:
                            curItem.NameFr = item.Name.ExtractText();
                            break;
                    }
                });
            }

            Serialize(Path.Join(OutputDir, "item_table.bin"), ItemTable.Serializer, itemTable);
        }

        private static void ExportMinionTable(Cyalume cyalume)
        {
            var minionTable = new MinionTable { Minions = new List<Minion>() };
            var languages = new[] { Language.English, Language.Japanese, Language.German, Language.French };
            foreach (var lang in languages)
            {
                var minionSheet = cyalume.GetExcelSheet<Lumina.Excel.Sheets.Companion>(lang);
                Parallel.ForEach(minionSheet, new ParallelOptions { MaxDegreeOfParallelism = 4 }, minion =>
                {
                    Minion curMinion;
                    lock (minionTable.Minions)
                    {
                        curMinion = minionTable.Minions.FirstOrDefault(m => m.Id == minion.RowId);
                        if (curMinion == null)
                        {
                            curMinion = new Minion { Id = minion.RowId, SortOrder = minion.Order };
                            minionTable.Minions.Add(curMinion);
                        }
                    }

                    switch (lang)
                    {
                        case Language.English:
                            curMinion.NameEn = minion.Singular.ExtractText();
                            break;
                        case Language.Japanese:
                            curMinion.NameJa = minion.Singular.ExtractText();
                            break;
                        case Language.German:
                            curMinion.NameDe = minion.Singular.ExtractText();
                            break;
                        case Language.French:
                            curMinion.NameFr = minion.Singular.ExtractText();
                            break;
                    }
                });
            }

            Serialize(Path.Join(OutputDir, "minion_table.bin"),MinionTable.Serializer, minionTable);
        }

        private static void ExportMountTable(Cyalume cyalume)
        {
            var mountTable = new MountTable { Mounts = new List<Mount>() };
            var languages = new[] { Language.English, Language.Japanese, Language.German, Language.French };
            foreach (var lang in languages)
            {
                var mountSheet = cyalume.GetExcelSheet<Lumina.Excel.Sheets.Mount>(lang);
                Parallel.ForEach(mountSheet, new ParallelOptions { MaxDegreeOfParallelism = 4 }, mount =>
                {
                    Mount curMount;
                    lock (mountTable.Mounts)
                    {
                        curMount = mountTable.Mounts.FirstOrDefault(m => m.Id == mount.RowId);
                        if (curMount == null)
                        {
                            curMount = new Mount { Id = mount.RowId, SortOrder = mount.UIPriority };
                            mountTable.Mounts.Add(curMount);
                        }
                    }

                    switch (lang)
                    {
                        case Language.English:
                            curMount.NameEn = mount.Singular.ExtractText();
                            break;
                        case Language.Japanese:
                            curMount.NameJa = mount.Singular.ExtractText();
                            break;
                        case Language.German:
                            curMount.NameDe = mount.Singular.ExtractText();
                            break;
                        case Language.French:
                            curMount.NameFr = mount.Singular.ExtractText();
                            break;
                    }
                });
            }

            Serialize(Path.Join(OutputDir, "mount_table.bin"),MountTable.Serializer, mountTable);
        }

        private static void ExportRaceTable(Cyalume cyalume)
        {
            var raceTable = new RaceTable { Races = new List<Race>() };
            var languages = new[] { Language.English, Language.Japanese, Language.German, Language.French };
            foreach (var lang in languages)
            {
                var raceSheet = cyalume.GetExcelSheet<Lumina.Excel.Sheets.Race>(lang);
                foreach (var race in raceSheet)
                {
                    var curRace = raceTable.Races.FirstOrDefault(r => r.Id == race.RowId);
                    if (curRace == null)
                    {
                        curRace = new Race { Id = race.RowId };
                        raceTable.Races.Add(curRace);
                    }

                    switch (lang)
                    {
                        case Language.English:
                            curRace.NameMasculineEn = race.Masculine.ExtractText();
                            curRace.NameFeminineEn = race.Feminine.ExtractText();
                            break;
                        case Language.Japanese:
                            curRace.NameMasculineJa = race.Masculine.ExtractText();
                            curRace.NameFeminineJa = race.Feminine.ExtractText();
                            break;
                        case Language.German:
                            curRace.NameMasculineDe = race.Masculine.ExtractText();
                            curRace.NameFeminineDe = race.Feminine.ExtractText();
                            break;
                        case Language.French:
                            curRace.NameMasculineFr = race.Masculine.ExtractText();
                            curRace.NameFeminineFr = race.Feminine.ExtractText();
                            break;
                    }
                }
            }

            Serialize(Path.Join(OutputDir, "race_table.bin"), RaceTable.Serializer, raceTable);
        }

        private static void ExportReputationTable(Cyalume cyalume)
        {
            var repTable = new ReputationTable { Reputations = new List<Reputation>() };
            var languages = new[] { Language.English, Language.Japanese, Language.German, Language.French };
            foreach (var lang in languages)
            {
                var repSheet = cyalume.GetExcelSheet<Lumina.Excel.Sheets.BeastReputationRank>(lang);
                foreach (var rep in repSheet)
                {
                    var curRep = repTable.Reputations.FirstOrDefault(r => r.Id == rep.RowId);
                    if (curRep == null)
                    {
                        curRep = new Reputation { Id = rep.RowId };
                        repTable.Reputations.Add(curRep);
                    }

                    switch (lang)
                    {
                        case Language.English:
                            curRep.NameEn = rep.Name.ExtractText();
                            break;
                        case Language.Japanese:
                            curRep.NameJa = rep.Name.ExtractText();
                            break;
                        case Language.German:
                            curRep.NameDe = rep.Name.ExtractText();
                            break;
                        case Language.French:
                            curRep.NameFr = rep.Name.ExtractText();
                            break;
                    }
                }
            }

            Serialize(Path.Join(OutputDir, "reputation_table.bin"), ReputationTable.Serializer, repTable);
        }

        private static void ExportTitleTable(Cyalume cyalume)
        {
            var titleTable = new TitleTable { Titles = new List<Title>() };
            var languages = new[] { Language.English, Language.Japanese, Language.German, Language.French };
            foreach (var lang in languages)
            {
                var titleSheet = cyalume.GetExcelSheet<Lumina.Excel.Sheets.Title>(lang);
                Parallel.ForEach(titleSheet, new ParallelOptions { MaxDegreeOfParallelism = 4 }, title =>
                {
                    Title curTitle;
                    lock (titleTable.Titles)
                    {
                        curTitle = titleTable.Titles.FirstOrDefault(t => t.Id == title.RowId);
                        if (curTitle == null)
                        {
                            curTitle = new Title { Id = title.RowId, IsPrefix = title.IsPrefix };
                            titleTable.Titles.Add(curTitle);
                        }
                    }

                    switch (lang)
                    {
                        case Language.English:
                            curTitle.NameMasculineEn = title.Masculine.ExtractText();
                            curTitle.NameFeminineEn = title.Feminine.ExtractText();
                            break;
                        case Language.Japanese:
                            curTitle.NameMasculineJa = title.Masculine.ExtractText();
                            curTitle.NameFeminineJa = title.Feminine.ExtractText();
                            break;
                        case Language.German:
                            curTitle.NameMasculineDe = title.Masculine.ExtractText();
                            curTitle.NameFeminineDe = title.Feminine.ExtractText();
                            break;
                        case Language.French:
                            curTitle.NameMasculineFr = title.Masculine.ExtractText();
                            curTitle.NameFeminineFr = title.Feminine.ExtractText();
                            break;
                    }
                });
            }

            Serialize(Path.Join(OutputDir, "title_table.bin"), TitleTable.Serializer, titleTable);
        }

        private static void ExportTownTable(Cyalume cyalume)
        {
            var townTable = new TownTable { Towns = new List<Town>() };
            var languages = new[] { Language.English, Language.Japanese, Language.German, Language.French };
            foreach (var lang in languages)
            {
                var townSheet = cyalume.GetExcelSheet<Lumina.Excel.Sheets.Town>(lang);
                foreach (var town in townSheet)
                {
                    var curTown = townTable.Towns.FirstOrDefault(t => t.Id == town.RowId);
                    if (curTown == null)
                    {
                        curTown = new Town { Id = town.RowId };
                        townTable.Towns.Add(curTown);
                    }

                    switch (lang)
                    {
                        case Language.English:
                            curTown.NameEn = town.Name.ExtractText();
                            break;
                        case Language.Japanese:
                            curTown.NameJa = town.Name.ExtractText();
                            break;
                        case Language.German:
                            curTown.NameDe = town.Name.ExtractText();
                            break;
                        case Language.French:
                            curTown.NameFr = town.Name.ExtractText();
                            break;
                    }
                }
            }

            Serialize(Path.Join(OutputDir, "town_table.bin"), TownTable.Serializer, townTable);
        }

        private static void ExportTribeTable(Cyalume cyalume)
        {
            var tribeTable = new TribeTable { Tribes = new List<Tribe>() };
            var languages = new[] { Language.English, Language.Japanese, Language.German, Language.French };
            foreach (var lang in languages)
            {
                var tribeSheet = cyalume.GetExcelSheet<Lumina.Excel.Sheets.Tribe>(lang);
                foreach (var tribe in tribeSheet)
                {
                    var curTribe = tribeTable.Tribes.FirstOrDefault(t => t.Id == tribe.RowId);
                    if (curTribe == null)
                    {
                        curTribe = new Tribe { Id = tribe.RowId };
                        tribeTable.Tribes.Add(curTribe);
                    }

                    switch (lang)
                    {
                        case Language.English:
                            curTribe.NameMasculineEn = tribe.Masculine.ExtractText();
                            curTribe.NameFeminineEn = tribe.Feminine.ExtractText();
                            break;
                        case Language.Japanese:
                            curTribe.NameMasculineJa = tribe.Masculine.ExtractText();
                            curTribe.NameFeminineJa = tribe.Feminine.ExtractText();
                            break;
                        case Language.German:
                            curTribe.NameMasculineDe = tribe.Masculine.ExtractText();
                            curTribe.NameFeminineDe = tribe.Feminine.ExtractText();
                            break;
                        case Language.French:
                            curTribe.NameMasculineFr = tribe.Masculine.ExtractText();
                            curTribe.NameFeminineFr = tribe.Feminine.ExtractText();
                            break;
                    }
                }
            }

            Serialize(Path.Join(OutputDir, "tribe_table.bin"), TribeTable.Serializer, tribeTable);
        }

        private static void Serialize<T>(string path, ISerializer<T> serializer, T obj) where T : class
        {
            var maxBytesNeeded = serializer.GetMaxSize(obj);
            var buffer = new byte[maxBytesNeeded];
            var bytesWritten = serializer.Write(buffer, obj);
            var bytesToWrite = buffer[..bytesWritten];
            File.WriteAllBytes(path, bytesToWrite);
        }
    }
}