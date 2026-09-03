using SharePointLinkManifestBuilder.Core.Models;
using SharePointLinkManifestBuilder.Core.Targets;

namespace SharePointLinkManifestBuilder.Core.Tests.Targets;

public class TargetPlannerTests
{
    [Fact]
    public void DetectOverlaps_DisjointFolders_FindsNone()
    {
        var targets = new[]
        {
            TestData.Target(startingPath: "Reports", recursive: true),
            TestData.Target(startingPath: "Archive", recursive: true),
        };

        Assert.Empty(TargetPlanner.DetectOverlaps(targets));
    }

    [Fact]
    public void DetectOverlaps_RecursiveParentAndChildFolder_FindsOverlap()
    {
        var parent = TestData.Target(startingPath: "Reports", recursive: true, targetId: "parent");
        var child = TestData.Target(startingPath: "Reports/Q1", recursive: false, targetId: "child");

        var overlaps = TargetPlanner.DetectOverlaps([parent, child]);

        var overlap = Assert.Single(overlaps);
        Assert.Equal("parent", overlap.Parent.TargetId);
        Assert.Equal("child", overlap.Child.TargetId);
        Assert.False(overlap.IsExactDuplicate);
    }

    /// <summary>
    /// The decisive rule. A non-recursive parent processes only its direct children, so a
    /// target on a subfolder does not overlap with it. Treating this as an overlap and dropping
    /// the child would silently skip every file the user asked for in that subfolder.
    /// </summary>
    [Fact]
    public void DetectOverlaps_NonRecursiveParentAndChildFolder_FindsNoOverlap()
    {
        var parent = TestData.Target(startingPath: "Reports", recursive: false);
        var child = TestData.Target(startingPath: "Reports/Q1", recursive: false);

        Assert.Empty(TargetPlanner.DetectOverlaps([parent, child]));
    }

    /// <summary>Prefix comparison must be segment-aware: "Reports" is not a parent of "ReportsArchive".</summary>
    [Fact]
    public void DetectOverlaps_SimilarlyNamedSiblingFolders_FindNoOverlap()
    {
        var a = TestData.Target(startingPath: "Reports", recursive: true);
        var b = TestData.Target(startingPath: "ReportsArchive", recursive: true);

        Assert.Empty(TargetPlanner.DetectOverlaps([a, b]));
    }

    [Fact]
    public void DetectOverlaps_DriveRootAndFolderInIt_FindsOverlapWhenRootRecurses()
    {
        var root = TestData.Target(startingPath: "", recursive: true, targetId: "root");
        var folder = TestData.Target(startingPath: "Reports", targetId: "folder");

        var overlap = Assert.Single(TargetPlanner.DetectOverlaps([root, folder]));
        Assert.Equal("root", overlap.Parent.TargetId);
    }

    [Fact]
    public void DetectOverlaps_SameFolderTwice_IsAnExactDuplicate()
    {
        var a = TestData.Target(startingPath: "Reports", targetId: "a");
        var b = TestData.Target(startingPath: "Reports", targetId: "b");

        var overlaps = TargetPlanner.DetectOverlaps([a, b]);

        var overlap = Assert.Single(overlaps);
        Assert.True(overlap.IsExactDuplicate);
    }

    [Fact]
    public void DetectOverlaps_DifferentDrives_FindNoOverlap()
    {
        var a = TestData.Target(driveId: TestData.DriveA, startingPath: "Reports", recursive: true);
        var b = TestData.Target(driveId: TestData.DriveB, startingPath: "Reports/Q1");

        Assert.Empty(TargetPlanner.DetectOverlaps([a, b]));
    }

    [Fact]
    public void DetectOverlaps_DifferentTenants_FindNoOverlap()
    {
        var a = TestData.Target(startingPath: "Reports", recursive: true);
        var b = a with { TargetId = "other", TenantId = "22222222-2222-2222-2222-222222222222" };

        Assert.Empty(TargetPlanner.DetectOverlaps([a, b]));
    }

    [Fact]
    public void DetectOverlaps_SiteTargetContainsItsLibraryRoot_EvenWhenNotRecursive()
    {
        var site = TestData.Target(TargetSourceType.SharePointSite, driveId: null, recursive: false, targetId: "site");
        var library = TestData.Target(TargetSourceType.DocumentLibrary, startingPath: "", targetId: "library");

        var overlap = Assert.Single(TargetPlanner.DetectOverlaps([site, library]));
        Assert.Equal("site", overlap.Parent.TargetId);
    }

    [Fact]
    public void DetectOverlaps_NonRecursiveSiteDoesNotContainASubfolder()
    {
        var site = TestData.Target(TargetSourceType.SharePointSite, driveId: null, recursive: false);
        var folder = TestData.Target(startingPath: "Reports/Q1");

        Assert.Empty(TargetPlanner.DetectOverlaps([site, folder]));
    }

    [Fact]
    public void DetectOverlaps_RecursiveSiteContainsASubfolder()
    {
        var site = TestData.Target(TargetSourceType.SharePointSite, driveId: null, recursive: true, targetId: "site");
        var folder = TestData.Target(startingPath: "Reports/Q1", targetId: "folder");

        var overlap = Assert.Single(TargetPlanner.DetectOverlaps([site, folder]));
        Assert.Equal("site", overlap.Parent.TargetId);
    }

    [Fact]
    public void Plan_KeepParent_DropsTheContainedTarget()
    {
        var parent = TestData.Target(startingPath: "Reports", recursive: true, targetId: "parent");
        var child = TestData.Target(startingPath: "Reports/Q1", targetId: "child");

        var plan = TargetPlanner.Plan([parent, child], OverlapResolution.KeepParent);

        Assert.Equal("parent", Assert.Single(plan.EffectiveTargets).TargetId);
        Assert.Single(plan.Removed);
        Assert.True(plan.HasOverlaps);
    }

    [Fact]
    public void Plan_KeepChild_DropsTheBroaderTarget()
    {
        var parent = TestData.Target(startingPath: "Reports", recursive: true, targetId: "parent");
        var child = TestData.Target(startingPath: "Reports/Q1", targetId: "child");

        var plan = TargetPlanner.Plan([parent, child], OverlapResolution.KeepChild);

        Assert.Equal("child", Assert.Single(plan.EffectiveTargets).TargetId);
        Assert.Contains(plan.Warnings, w => w.Contains("will not be processed", StringComparison.Ordinal));
    }

    [Fact]
    public void Plan_KeepBothDeduplicate_RetainsBothTargets()
    {
        var parent = TestData.Target(startingPath: "Reports", recursive: true, targetId: "parent");
        var child = TestData.Target(startingPath: "Reports/Q1", targetId: "child");

        var plan = TargetPlanner.Plan([parent, child], OverlapResolution.KeepBothDeduplicate);

        Assert.Equal(2, plan.EffectiveTargets.Count);
        Assert.Contains(plan.Warnings, w => w.Contains("processed once", StringComparison.Ordinal));
    }

    /// <summary>An exact duplicate is collapsed regardless of the chosen resolution.</summary>
    [Fact]
    public void Plan_ExactDuplicate_IsCollapsedEvenWhenKeepingBoth()
    {
        var a = TestData.Target(startingPath: "Reports", targetId: "a");
        var b = TestData.Target(startingPath: "Reports", targetId: "b");

        var plan = TargetPlanner.Plan([a, b], OverlapResolution.KeepBothDeduplicate);

        Assert.Single(plan.EffectiveTargets);
    }

    [Fact]
    public void Plan_DisabledTargets_AreExcluded()
    {
        var enabled = TestData.Target(startingPath: "Reports", targetId: "on");
        var disabled = TestData.Target(startingPath: "Archive", targetId: "off") with { IsEnabled = false };

        var plan = TargetPlanner.Plan([enabled, disabled]);

        Assert.Equal("on", Assert.Single(plan.EffectiveTargets).TargetId);
    }

    [Theory]
    [InlineData("Reports/Q1", "Reports", true)]
    [InlineData("Reports/Q1/Jan", "Reports", true)]
    [InlineData("ReportsArchive", "Reports", false)]
    [InlineData("Reports", "Reports", false)]
    [InlineData("Reports", "", true)]
    [InlineData("", "", false)]
    public void IsUnder_ComparesWholeSegments(string candidate, string ancestor, bool expected) =>
        Assert.Equal(expected, TargetPlanner.IsUnder(candidate, ancestor));

    [Theory]
    [InlineData(null, "")]
    [InlineData("", "")]
    [InlineData("/Reports/", "Reports")]
    [InlineData("\\Reports\\Q1", "Reports/Q1")]
    public void NormalizePath_ProducesForwardSlashedTrimmedPaths(string? input, string expected) =>
        Assert.Equal(expected, TargetPlanner.NormalizePath(input));
}
