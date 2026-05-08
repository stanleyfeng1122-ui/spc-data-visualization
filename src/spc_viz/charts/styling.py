"""Plotly figure styling helpers.

The ``finalize_plotly_style`` function applies the consistent white
background, font, and axis colour treatment used across every chart in
the app. Pure code movement from the original ``chart_utils`` module.
"""


# ---------------------------------------------------------------------------
# Plotly white-background finalizer
# ---------------------------------------------------------------------------


def finalize_plotly_style(fig):
    """Apply consistent white background and black text to a Plotly figure."""
    fig.update_layout(
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(color="#000000", family="SF Pro Display, SF Pro, -apple-system, sans-serif"),
        title=dict(font=dict(color="#000000")),
        legend=dict(font=dict(color="#000000"), title=dict(font=dict(color="#000000"))),
        xaxis=dict(
            tickfont=dict(color="#000000"), title=dict(font=dict(color="#000000")), color="#000000"
        ),
        yaxis=dict(
            tickfont=dict(color="#000000"), title=dict(font=dict(color="#000000")), color="#000000"
        ),
    )
    return fig
