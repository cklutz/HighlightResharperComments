using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Windows.Media;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Text.Tagging;
using Microsoft.VisualStudio.Utilities;

namespace HighlightReSharperComments
{
    [Export(typeof(IViewTaggerProvider))]
    [ContentType("any")]
    [TagType(typeof(ClassificationTag))]
    public sealed class ReSharperCommentTaggerProvider : IViewTaggerProvider
    {
        [Import]
        public IClassificationTypeRegistryService Registry;

        [Import]
        internal ITextSearchService TextSearchService { get; set; }

        public ITagger<T> CreateTagger<T>(ITextView textView, ITextBuffer buffer) where T : ITag
        {
            if (buffer != textView.TextBuffer)
                return null;

            var classType = Registry.GetClassificationType("resharper-comment");
            return new ReSharperCommentTagger(textView, TextSearchService, classType) as ITagger<T>;
        }
    }

    public static class TypeExports
    {
        [Export(typeof(ClassificationTypeDefinition))]
        [Name("resharper-comment")]
        public static ClassificationTypeDefinition OrdinaryClassificationType;
    }

    [Export(typeof(EditorFormatDefinition))]
    [ClassificationType(ClassificationTypeNames = "resharper-comment")]
    [Name("resharper-comment")]
    [UserVisible(true)]
    [Order(After = Priority.High)]
    public sealed class RegionForeground : ClassificationFormatDefinition
    {
        public RegionForeground()
        {
            DisplayName = "Highlight ReSharper Comments";
            ForegroundColor = Colors.DimGray;
        }
    }

    public sealed class ReSharperCommentTagger : ITagger<ClassificationTag>
    {
        private readonly ITextView m_view;
        private readonly ITextSearchService m_searchService;
        private readonly IClassificationType m_type;
        private NormalizedSnapshotSpanCollection m_currentSpans;

        public event EventHandler<SnapshotSpanEventArgs> TagsChanged = delegate { };

        public ReSharperCommentTagger(ITextView view, ITextSearchService searchService, IClassificationType type)
        {
            m_view = view;
            m_searchService = searchService;
            m_type = type;

            m_currentSpans = GetWordSpans(m_view.TextSnapshot);

            m_view.GotAggregateFocus += SetupSelectionChangedListener;
        }

        private void SetupSelectionChangedListener(object sender, EventArgs e)
        {
            if (m_view != null)
            {
                m_view.LayoutChanged += ViewLayoutChanged;
                m_view.GotAggregateFocus -= SetupSelectionChangedListener;
            }
        }

        private void ViewLayoutChanged(object sender, TextViewLayoutChangedEventArgs e)
        {
            if (e.OldSnapshot != e.NewSnapshot)
            {
                m_currentSpans = GetWordSpans(e.NewSnapshot);
                TagsChanged(this, new SnapshotSpanEventArgs(new SnapshotSpan(e.NewSnapshot, 0, e.NewSnapshot.Length)));
            }
        }

        private NormalizedSnapshotSpanCollection GetWordSpans(ITextSnapshot snapshot)
        {
            var wordSpans = new List<SnapshotSpan>();
            wordSpans.AddRange(FindAll(@"// ReSharper disable", snapshot).Select(regionLine => regionLine.Start.GetContainingLine().Extent));
            wordSpans.AddRange(FindAll(@"// ReSharper restore", snapshot).Select(regionLine => regionLine.Start.GetContainingLine().Extent));
            return new NormalizedSnapshotSpanCollection(wordSpans);
        }

        private IEnumerable<SnapshotSpan> FindAll(String searchPattern, ITextSnapshot textSnapshot)
        {
            if (textSnapshot == null)
                return null;

            return m_searchService.FindAll(
                new FindData(searchPattern, textSnapshot)
                {
                    FindOptions = FindOptions.WholeWord | FindOptions.MatchCase
                });
        }

        public IEnumerable<ITagSpan<ClassificationTag>> GetTags(NormalizedSnapshotSpanCollection spans)
        {
            if (spans == null || spans.Count == 0 || m_currentSpans.Count == 0)
                yield break;

            ITextSnapshot snapshot = m_currentSpans[0].Snapshot;
            spans = new NormalizedSnapshotSpanCollection(spans.Select(s => s.TranslateTo(snapshot, SpanTrackingMode.EdgeExclusive)));

            foreach (var span in NormalizedSnapshotSpanCollection.Intersection(m_currentSpans, spans))
            {
                yield return new TagSpan<ClassificationTag>(span, new ClassificationTag(m_type));
            }
        }
    }
}