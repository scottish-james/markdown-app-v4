"""
Diagram Analyzer - v19 Scoring System for PowerPoint Diagram Detection
Analyzes slide content to identify potential diagrams with probability scoring
"""


class DiagramAnalyzer:
    """
    Analyzes structured slide data to identify potential diagrams using
    the v19 scoring system with sophisticated rules and probability calculation.
    """

    def analyze_structured_data_for_diagrams(self, structured_data):
        """
        Analyze extracted structured data for diagram presence.

        Args:
            structured_data (dict): Structured presentation data

        Returns:
            str: Diagram analysis summary or None if no diagrams found
        """
        try:
            diagram_slides = []

            for slide_idx, slide in enumerate(structured_data["slides"]):
                score_analysis = self.score_slide_for_diagram(slide)
                if score_analysis["probability"] >= 40:  # 40%+ probability threshold
                    diagram_slides.append({
                        "slide": slide_idx + 1,
                        "analysis": score_analysis
                    })

            # Generate detailed summary
            if diagram_slides:
                summary = "## DIAGRAM ANALYSIS (v19 Scoring System)\n\n"
                summary += "**Slides with potential diagrams:**\n\n"

                for slide_info in diagram_slides:
                    analysis = slide_info["analysis"]
                    summary += f"- **Slide {slide_info['slide']}**: {analysis['probability']}% probability "
                    summary += f"(Score: {analysis['total_score']}) - {', '.join(analysis['reasons'])}\n"
                    summary += f"  - Shapes: {analysis['shape_count']}, Lines: {analysis['line_count']}, Arrows: {analysis['arrow_count']}\n\n"

                return summary

            return None

        except Exception as e:
            return f"\n\n<!-- v19 Diagram analysis error: {e} -->"

    def score_slide_for_diagram(self, slide_data):
        """
        Score a slide for diagram probability using v19 sophisticated rules.

        Args:
            slide_data (dict): Slide data with content blocks

        Returns:
            dict: Scoring analysis with probability and reasoning
        """
        content_blocks = slide_data.get("content_blocks", [])

        # Collect and categorize shapes and elements
        shapes, lines, arrows, text_blocks = self._categorize_slide_elements(content_blocks)

        # Calculate score using v19 rules
        score = 0
        reasons = []

        # Rule 1: Arrow/Line threshold (20+ points each)
        if len(arrows) > 0:
            score += 20
            reasons.append(f"block_arrows:{len(arrows)}")

        if len(lines) >= 3:
            score += 20
            reasons.append(f"connector_lines:{len(lines)}")

        # Rule 2: Line-to-shape ratio (15 points)
        total_lines = len(lines) + len(arrows)
        if len(shapes) > 0:
            line_ratio = total_lines / len(shapes)
            if line_ratio >= 0.5:
                score += 15
                reasons.append(f"line_ratio:{line_ratio:.1f}")

        # Rule 3: Spatial layout analysis (10-15 points)
        layout_score = self._analyze_spatial_layout(shapes)
        score += layout_score["score"]
        if layout_score["score"] > 0:
            reasons.append(f"layout:{layout_score['type']}")

        # Rule 4: Shape variety (10-15 points)
        variety_score = self._analyze_shape_variety(shapes)
        score += variety_score
        if variety_score > 0:
            reasons.append(f"variety:{variety_score}")

        # Rule 5: Text density analysis (10 points)
        text_score = self._analyze_text_density(text_blocks)
        score += text_score
        if text_score > 0:
            reasons.append(f"short_text:{text_score}")

        # Rule 6: Flow patterns (20 points)
        flow_score = self._analyze_flow_patterns(shapes, lines, arrows, text_blocks)
        score += flow_score
        if flow_score > 0:
            reasons.append(f"flow_pattern:{flow_score}")

        # Negative indicators
        negative_score = self._analyze_negative_indicators(text_blocks, shapes)
        score += negative_score  # negative_score will be negative or 0
        if negative_score < 0:
            reasons.append(f"negatives:{negative_score}")

        # Convert score to probability
        probability = self._calculate_probability_from_score(score)

        return {
            "total_score": score,
            "probability": probability,
            "reasons": reasons,
            "shape_count": len(shapes),
            "line_count": len(lines),
            "arrow_count": len(arrows)
        }

    def _categorize_slide_elements(self, content_blocks):
        """
        Categorize slide elements into shapes, lines, arrows, and text blocks.

        Args:
            content_blocks (list): List of content blocks from slide

        Returns:
            tuple: (shapes, lines, arrows, text_blocks)
        """
        shapes = []
        lines = []
        arrows = []
        text_blocks = []

        for block in content_blocks:
            if block.get("type") == "line":
                lines.append(block)
            elif block.get("type") == "arrow":
                arrows.append(block)
            elif block.get("type") == "text":
                text_blocks.append(block)
                shapes.append(block)
            elif block.get("type") in ["shape", "image", "chart"]:
                shapes.append(block)
            elif block.get("type") == "group":
                # Recursively analyze group contents
                group_analysis = self._analyze_group_contents(block)
                shapes.extend(group_analysis["shapes"])
                lines.extend(group_analysis["lines"])
                arrows.extend(group_analysis["arrows"])
                text_blocks.extend(group_analysis["text_blocks"])

        return shapes, lines, arrows, text_blocks

    def _analyze_group_contents(self, group_block):
        """
        Recursively analyze group contents for diagram elements.

        Args:
            group_block (dict): Group content block

        Returns:
            dict: Categorized elements from group
        """
        result = {"shapes": [], "lines": [], "arrows": [], "text_blocks": []}

        for extracted_block in group_block.get("extracted_blocks", []):
            if extracted_block.get("type") == "line":
                result["lines"].append(extracted_block)
            elif extracted_block.get("type") == "arrow":
                result["arrows"].append(extracted_block)
            elif extracted_block.get("type") == "text":
                result["text_blocks"].append(extracted_block)
                result["shapes"].append(extracted_block)
            elif extracted_block.get("type") in ["shape", "image", "chart"]:
                result["shapes"].append(extracted_block)

        return result

    def _analyze_spatial_layout(self, shapes):
        """
        Analyze spatial layout patterns in shapes.

        Args:
            shapes (list): List of shape elements

        Returns:
            dict: Layout analysis with score and type
        """
        if len(shapes) < 3:
            return {"score": 0, "type": "insufficient"}

        positions = []
        for shape in shapes:
            pos = shape.get("position")
            if pos:
                positions.append((pos["top"], pos["left"]))

        if len(positions) < 3:
            return {"score": 0, "type": "no_position_data"}

        # Calculate spatial distribution
        tops = [p[0] for p in positions]
        lefts = [p[1] for p in positions]

        top_range = max(tops) - min(tops) if tops else 0
        left_range = max(lefts) - min(lefts) if lefts else 0

        # Check for grid-like arrangement
        unique_tops = len(set(round(t / 100000) for t in tops))  # Group by approximate position
        unique_lefts = len(set(round(l / 100000) for l in lefts))

        if unique_tops >= 2 and unique_lefts >= 2:
            return {"score": 15, "type": "grid_layout"}
        elif top_range > 1000000 and left_range > 1000000:
            return {"score": 10, "type": "spread_layout"}
        else:
            return {"score": 0, "type": "linear_layout"}

    def _analyze_shape_variety(self, shapes):
        """
        Analyze variety in shape types and sizes.

        Args:
            shapes (list): List of shape elements

        Returns:
            int: Variety score
        """
        if len(shapes) < 2:
            return 0

        shape_types = set()
        sizes = []

        for shape in shapes:
            shape_types.add(shape.get("type", "unknown"))
            pos = shape.get("position")
            if pos:
                size = pos["width"] * pos["height"]
                sizes.append(size)

        score = 0

        # Multiple shape types
        if len(shape_types) >= 3:
            score += 15
        elif len(shape_types) >= 2:
            score += 10

        # Consistent sizing (indicates process flow)
        if len(sizes) >= 3:
            avg_size = sum(sizes) / len(sizes)
            variations = [abs(size - avg_size) / avg_size for size in sizes if avg_size > 0]
            if variations and max(variations) < 0.5:  # Less than 50% variation
                score += 5

        return score

    def _analyze_text_density(self, text_blocks):
        """
        Analyze text characteristics for diagram indicators.

        Args:
            text_blocks (list): List of text blocks

        Returns:
            int: Text density score
        """
        if not text_blocks:
            return 0

        short_text_count = 0
        total_blocks = len(text_blocks)

        for block in text_blocks:
            # Count average words per paragraph
            total_words = 0
            para_count = 0

            for para in block.get("paragraphs", []):
                clean_text = para.get("clean_text", "")
                if clean_text:
                    words = len(clean_text.split())
                    total_words += words
                    para_count += 1

            if para_count > 0:
                avg_words = total_words / para_count
                if avg_words <= 5:  # Short labels
                    short_text_count += 1

        # Score based on percentage of short text blocks
        if total_blocks > 0:
            short_ratio = short_text_count / total_blocks
            if short_ratio >= 0.7:  # 70%+ short text
                return 10
            elif short_ratio >= 0.5:  # 50%+ short text
                return 5

        return 0

    def _analyze_flow_patterns(self, shapes, lines, arrows, text_blocks):
        """
        Analyze for flow patterns and process keywords.

        Args:
            shapes (list): Shape elements
            lines (list): Line elements
            arrows (list): Arrow elements
            text_blocks (list): Text elements

        Returns:
            int: Flow pattern score
        """
        score = 0

        # Check for workflow keywords
        flow_keywords = ["start", "begin", "end", "finish", "process", "step", "decision"]
        action_words = ["create", "update", "check", "verify", "send", "receive", "analyze"]

        all_text = ""
        for block in text_blocks:
            for para in block.get("paragraphs", []):
                all_text += " " + para.get("clean_text", "").lower()

        flow_matches = sum(1 for keyword in flow_keywords if keyword in all_text)
        action_matches = sum(1 for keyword in action_words if keyword in all_text)

        if flow_matches >= 2:
            score += 20
        elif flow_matches >= 1:
            score += 10

        if action_matches >= 3:
            score += 10

        # Bonus for having both shapes and connecting elements
        if len(shapes) >= 3 and (len(lines) > 0 or len(arrows) > 0):
            score += 15

        return score

    def _analyze_negative_indicators(self, text_blocks, shapes):
        """
        Check for negative indicators that suggest NOT a diagram.

        Args:
            text_blocks (list): Text elements
            shapes (list): Shape elements

        Returns:
            int: Negative score (0 or negative)
        """
        score = 0

        # Check for long paragraphs
        long_text_count = 0
        bullet_count = 0

        for block in text_blocks:
            for para in block.get("paragraphs", []):
                clean_text = para.get("clean_text", "")
                if clean_text:
                    word_count = len(clean_text.split())
                    if word_count > 20:  # Long paragraph
                        long_text_count += 1

                    # Check for bullet points
                    if para.get("hints", {}).get("is_bullet", False):
                        bullet_count += 1

        # Penalize long text
        if long_text_count >= 2:
            score -= 15

        # Penalize if mostly bullet points
        total_paras = sum(len(block.get("paragraphs", [])) for block in text_blocks)
        if total_paras > 0 and bullet_count / total_paras > 0.8:
            score -= 10

        # Penalize single column layout (all shapes vertically aligned)
        if len(shapes) >= 3:
            positions = [s.get("position") for s in shapes if s.get("position")]
            if len(positions) >= 3:
                lefts = [p["left"] for p in positions]
                left_variance = max(lefts) - min(lefts) if lefts else 0
                if left_variance < 500000:  # Very narrow horizontal spread
                    score -= 10

        return score

    def _calculate_probability_from_score(self, score):
        """
        Convert raw score to probability percentage.

        Args:
            score (int): Raw diagram score

        Returns:
            int: Probability percentage (0-100)
        """
        if score >= 60:
            return 95
        elif score >= 40:
            return 75
        elif score >= 20:
            return 40
        else:
            return 10

    