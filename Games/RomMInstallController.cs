using Playnite.SDK;
using Playnite.SDK.Models;
using Playnite.SDK.Plugins;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading;
using SharpCompress.Archives;
using SharpCompress.Common;
using System.Linq;

namespace RomM.Games
{
    internal class RomMInstallController : InstallController
    {
        protected readonly IRomM _romM;
        protected CancellationTokenSource _watcherToken;
        public ILogger Logger => LogManager.GetLogger();

        internal RomMInstallController(Game game, IRomM romM) : base(game)
        {
            Name = "Download";
            _romM = romM;
        }

        public override void Dispose()
        {
            try { _watcherToken?.Cancel(); } catch { }
            base.Dispose();
        }

        public override void Install(InstallActionArgs args)
        {
            var info = Game.GetRomMGameInfo();

            var dstPath = info.Mapping?.DestinationPathResolved
                ?? throw new Exception("Mapped emulator data cannot be found, try removing and re-adding.");

            _watcherToken = new CancellationTokenSource();

            _romM.Playnite.Dialogs.ActivateGlobalProgress(
                async progress =>
                {
                    try
                    {
                        progress.ProgressMaxValue = 100;
                        progress.IsIndeterminate = true;
                        progress.Text = $"Starting download: {Game.Name}";

                        HttpResponseMessage response =
                            await RomM.GetAsync(info.DownloadUrl, HttpCompletionOption.ResponseHeadersRead);

                        response.EnsureSuccessStatusCode();
                        var totalBytes = response.Content.Headers.ContentLength;

                        // paths
                        string installDir = Path.Combine(dstPath, Path.GetFileNameWithoutExtension(info.FileName));
                        string gamePath = info.HasMultipleFiles
                            ? Path.Combine(installDir, info.FileName + ".zip")
                            : Path.Combine(installDir, info.FileName);

                        Directory.CreateDirectory(installDir);

                        progress.IsIndeterminate = !totalBytes.HasValue;
                        progress.Text = "Downloading...";

                        long downloaded = 0;
                        byte[] buffer = new byte[1024 * 256];

                        using (var httpStream = await response.Content.ReadAsStreamAsync())
                        using (var fileStream = new FileStream(
                            gamePath,
                            FileMode.Create,
                            FileAccess.Write,
                            FileShare.None,
                            buffer.Length,
                            true))
                        {
                            while (true)
                            {
                                progress.CancelToken.ThrowIfCancellationRequested();

                                int read = await httpStream.ReadAsync(buffer, 0, buffer.Length);
                                if (read <= 0)
                                    break;

                                await fileStream.WriteAsync(buffer, 0, read);
                                downloaded += read;

                                if (totalBytes.HasValue && totalBytes.Value > 0)
                                {
                                    double pct = (double)downloaded / totalBytes.Value;
                                    progress.CurrentProgressValue = Math.Min(85, pct * 85);
                                    progress.Text = $"Downloading {Game.Name}... {pct * 100:0}%";
                                }
                            }
                        }

                        progress.IsIndeterminate = false;
                        progress.CurrentProgressValue = 85;
                        progress.Text = "Extracting...";

                        if (info.HasMultipleFiles || (info.Mapping.AutoExtract && IsFileCompressed(gamePath)))
                        {
                            ExtractArchiveWithProgress(gamePath, installDir, progress);
                            File.Delete(gamePath);
                        }

                        progress.CurrentProgressValue = 100;
                        progress.Text = "Installed";

                        var game = _romM.Playnite.Database.Games[Game.Id];
                        game.IsInstalled = true;
                        _romM.Playnite.Database.Games.Update(game);

                        InvokeOnInstalled(new GameInstalledEventArgs(new GameInstallationData
                        {
                            InstallDirectory = installDir,
                            Roms = new List<GameRom> { new GameRom(Game.Name, gamePath) }
                        }));
                    }
                    catch (OperationCanceledException)
                    {
                        InvokeOnInstallationCancelled(new GameInstallationCancelledEventArgs());
                    }
                },
                new GlobalProgressOptions($"Installing {Game.Name}", true)
            );
        }

        private static string[] GetRomFiles(string installDir, List<string> supportedFileTypes)
        {
            if (installDir == null || installDir.Contains("../") || installDir.Contains(@"..\"))
            {
                throw new ArgumentException("Invalid file path");
            }

            if (supportedFileTypes == null || supportedFileTypes.Count == 0)
            {
                return Directory.GetFiles(installDir, "*", SearchOption.AllDirectories)
                    .Where(file => !file.EndsWith(".m3u", StringComparison.OrdinalIgnoreCase))
                    .ToArray();
            }

            return supportedFileTypes.SelectMany(fileType =>
            {
                if (fileType == null || fileType.Contains("../") || fileType.Contains(@"..\"))
                {
                    throw new ArgumentException("Invalid file path");
                }

                return Directory.GetFiles(installDir, "*." + fileType, SearchOption.AllDirectories)
                    .Where(file => !file.Contains("../") && !file.Contains(@"..\"));
            }).ToArray();
        }

        private static List<string> GetEmulatorSupportedFileTypes(RomMGameInfo info)
        {
            if (info.Mapping.EmulatorProfile is CustomEmulatorProfile customProfile)
            {
                return customProfile.ImageExtensions;
            }
            else if (info.Mapping.EmulatorProfile is BuiltInEmulatorProfile builtInProfile)
            {
                return API.Instance.Emulation.Emulators
                    .FirstOrDefault(e => e.Id == info.Mapping.Emulator.BuiltInConfigId)?
                    .Profiles
                    .FirstOrDefault(p => p.Name == builtInProfile.Name)?
                    .ImageExtensions;
            }

            return null;
        }

        private static bool IsFileCompressed(string filePath)
        {
            if (Path.GetExtension(filePath).Equals(".iso", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return ArchiveFactory.IsArchive(filePath, out var type);
        }

        private void ExtractArchiveWithProgress(
            string archivePath,
            string installDir,
            GlobalProgressActionArgs progress)
        {
            using (var archive = ArchiveFactory.Open(archivePath))
            {
                var entries = archive.Entries.Where(e => !e.IsDirectory).ToList();
                int total = entries.Count;
                int done = 0;

                foreach (var entry in entries)
                {
                    progress.CancelToken.ThrowIfCancellationRequested();

                    entry.WriteToDirectory(installDir, new ExtractionOptions
                    {
                        ExtractFullPath = true,
                        Overwrite = true
                    });

                    done++;
                    double pct = total > 0 ? (double)done / total : 1.0;
                    progress.CurrentProgressValue = 85 + pct * 15;
                    progress.Text = $"Extracting {Game.Name}... {pct * 100:0}%";
                }
            }
        }

        private void ExtractNestedArchives(string directoryPath, CancellationToken ct, GlobalProgressActionArgs progress)
        {
            if (directoryPath == null || directoryPath.Contains("../") || directoryPath.Contains(@"..\"))
            {
                throw new ArgumentException("Invalid file path");
            }

            foreach (var file in Directory.GetFiles(directoryPath))
            {
                progress.CancelToken.ThrowIfCancellationRequested();

                if (IsFileCompressed(file))
                {
                    progress.Text = $"Extracting nested: {Path.GetFileName(file)}";
                    ExtractArchiveWithProgress(file, directoryPath, progress);
                    try { File.Delete(file); } catch { }
                }
            }
        }
    }
}
