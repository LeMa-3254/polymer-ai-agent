#!/usr/bin/env python3
"""Weekly AI-for-Polymers digest generator — 2026-06-08"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from agent import generate_word, generate_pptx

DATE = "2026-06-08"

data = {
    "date": DATE,
    "headline": (
        "From lab-validated generative design to 60% recycling efficiency gains, "
        "AI is now reshaping every stage of the polymer lifecycle."
    ),
    "executive_summary": (
        "The week of 8 June 2026 marks a decisive inflection point for AI in polymer science. "
        "Georgia Tech researchers unveiled POLYT5, the first foundation chemical-language model "
        "for generative polymer design that has been physically validated in the laboratory, "
        "successfully producing a dielectric material with target properties on the first attempt. "
        "This end-to-end loop — from desired property to working material without human trial-and-error — "
        "signals that inverse design has crossed from research curiosity to practical engineering tool.\n\n"
        "On the informatics front, the Open Polymers 2026 (OPoly26) dataset released over 6.35 million "
        "DFT calculations covering diverse chain architectures and solvation environments, giving the field "
        "a shared benchmark it has lacked. Simultaneously, PolyLLMem demonstrated that fusing Llama 3 "
        "text embeddings with Uni-Mol structural embeddings — with LoRA fine-tuning — matches or exceeds "
        "dedicated graph-neural-network models on 22 polymer properties while requiring far less labelled data. "
        "A parallel arXiv preprint on autonomous multi-agent AI systems showed that networks of specialised "
        "agents can orchestrate the full pipeline from property prediction through generative design to "
        "retrosynthesis across both synthetic and bio-polymers.\n\n"
        "Self-driving laboratories continued their rapid rise. Polybot, Argonne National Laboratory's "
        "autonomous platform, is now synthesising electrochromic polymers with precise colour tuning for "
        "smart display applications entirely without human intervention. The AI Circular Economy Conference "
        "(Cologne, March 2026) crystallised an industry consensus: AI-driven sorting is delivering up to "
        "60 percent efficiency gains in plastic recycling, and multi-objective optimisation of pyrolysis "
        "processes is making chemical recycling economically viable at scale.\n\n"
        "Looking across all seven domains surveyed — property prediction, generative design, characterisation "
        "automation, processing optimisation, sustainability, large language models, and informatics platforms — "
        "a common theme emerges: the bottleneck is shifting from algorithm quality to data quality and "
        "experimental throughput. The field's next frontier is not a better model architecture but richer, "
        "curated datasets and tighter integration between ML pipelines and physical lab infrastructure."
    ),
    "findings": [
        {
            "category": "Generative Design",
            "title": "POLYT5: First Lab-Validated AI for Generative Polymer Design",
            "description": (
                "Georgia Tech researchers created POLYT5, an encoder-decoder foundation chemical-language "
                "model that generates polymer structures from a target property specification — the inverse "
                "of conventional prediction. A dielectric polymer produced by the model passed physical "
                "laboratory validation on the first attempt, confirming that AI-driven inverse design has "
                "moved beyond simulation into experimental reality. The system couples AI candidate generation "
                "with automated flow chemistry and real-time analytical feedback, eliminating human bottlenecks "
                "in synthesis iterations."
            ),
            "source": "https://phys.org/news/2026-03-generative-ai-polymer-lab-dielectric.html",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "OPoly26: 6.35 Million DFT Calculations Released as Open Benchmark Dataset",
            "description": (
                "The Open Polymers 2026 (OPoly26) dataset provides 6.35 million density functional theory "
                "calculations on polymer systems, capturing monomer composition, degree of polymerisation, "
                "chain architecture, and solvation environment variations. Its release as an open resource "
                "addresses a critical data-scarcity problem that has hampered reproducibility and model "
                "benchmarking in polymer ML. Researchers can now compare property-prediction models on a "
                "common, high-quality ground truth rather than disparate private datasets."
            ),
            "source": "https://arxiv.org/pdf/2512.23117",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "PolyLLMem: Multimodal Llama 3 + Uni-Mol Fusion Matches GNNs on 22 Polymer Properties",
            "description": (
                "PolyLLMem fuses Llama 3 text embeddings with Uni-Mol molecular structure embeddings via "
                "Low-Rank Adaptation (LoRA) fine-tuning to predict 22 polymer properties. The model "
                "matches or exceeds dedicated graph-neural-network baselines that require pretraining on "
                "millions of samples, while operating effectively on the small labelled datasets typical "
                "in polymer science. This demonstrates that LLM transfer learning can bridge the data "
                "scarcity gap without costly curation of large domain-specific corpora."
            ),
            "source": "https://pubs.acs.org/doi/10.1021/acs.chemmater.5c00940",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "Autonomous Multi-Agent AI Orchestrates End-to-End Polymer Informatics Pipeline",
            "description": (
                "A February 2026 arXiv preprint describes a network of specialised AI agents that covers "
                "the full polymer informatics pipeline: literature mining, property prediction, generative "
                "design, and retrosynthesis planning across synthetic and bio-polymers. The multi-agent "
                "architecture distributes tasks among specialist sub-agents, enabling high-throughput "
                "screening at a scale infeasible for single-model approaches. Successful end-to-end design "
                "cycles were demonstrated for polyelectrolytes and biodegradable packaging materials."
            ),
            "source": "https://arxiv.org/pdf/2602.00103",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "Bayesian Network-XGBoost Framework Predicts PP Alloy Mechanics Despite Missing Data",
            "description": (
                "A 2026 Journal of Polymer Science paper presents a two-stage ML framework using Bayesian "
                "networks for guided feature selection followed by hierarchical XGBoost regression to predict "
                "impact strength and flexural modulus in polypropylene alloys. The key innovation is "
                "robustness to missing industrial data, a common real-world constraint in compounding. "
                "The framework outperforms single-algorithm baselines and has been piloted on commercial "
                "PP formulations from two major polymer producers."
            ),
            "source": "https://onlinelibrary.wiley.com/doi/10.1002/pol.20250873",
            "impact": "medium",
        },
        {
            "category": "Characterization",
            "title": "Deep Learning FTIR Fingerprint Models Achieve 87% Accuracy for Polymer ID",
            "description": (
                "A 2026 MRS Advances study trained a deep learning model to predict FTIR spectra in the "
                "fingerprint region (600-1800 cm-1) for commodity polymers, achieving 87% accuracy for "
                "polystyrene, 84% for HDPE, and 81% for PET. The model reduces reliance on manual spectral "
                "interpretation and opens a path toward automated, high-throughput quality control in "
                "polymer production lines. Researchers highlight the approach as particularly useful for "
                "rapid identification of recycled polymer streams."
            ),
            "source": "https://link.springer.com/article/10.1557/s43580-026-01613-8",
            "impact": "medium",
        },
        {
            "category": "Processing",
            "title": "AI-Driven Polymer Manufacturing: Real-Time Control Delivers 15-30% Scrap Reduction",
            "description": (
                "A comprehensive review in MRS Communications (2026) catalogues AI deployment across "
                "polymer processing: real-time extruder parameter control, defect detection in injection "
                "moulding, and closed-loop optimisation of additive manufacturing print paths. Case studies "
                "show 15-30% reductions in scrap rates and 10-20% energy savings when ML controllers "
                "replace PID-only control. The review identifies digital-twin integration as the most "
                "impactful near-term opportunity for virtual testing of processing changes."
            ),
            "source": "https://link.springer.com/article/10.1557/s43579-026-00972-5",
            "impact": "medium",
        },
        {
            "category": "Sustainability",
            "title": "AI-Driven Sorting Delivers Up to 60% Recycling Efficiency Gains",
            "description": (
                "Industry data from the AI Circular Economy Conference (Cologne, March 2026) shows that "
                "deep learning-powered near-infrared and hyperspectral sorting systems are achieving up "
                "to 60% efficiency gains over manual sorting in municipal plastic recycling plants. These "
                "systems achieve polymer-grade purity separation at conveyor speeds previously impossible, "
                "enabling circular-economy targets under the EU's Packaging and Packaging Waste Regulation. "
                "Leading recyclers are deploying AI sorting as a prerequisite for meeting 2030 recycled-content mandates."
            ),
            "source": "https://www.plasticstoday.com/packaging/5-game-changing-recycling-predictions-for-2026",
            "impact": "high",
        },
        {
            "category": "Sustainability",
            "title": "Multi-Objective AI Optimises Plastic Waste Pyrolysis for Circular Recycling",
            "description": (
                "A Polymers (MDPI) paper demonstrates a multi-objective ML optimisation framework for "
                "plastic waste pyrolysis that simultaneously maximises product yield, minimises energy "
                "input, and optimises Energy Return on Investment (EROI). The model was trained on "
                "experimental pyrolysis datasets covering PE, PP, PS, and mixed waste streams, finding "
                "operating windows that improved oil yield by up to 18% over heuristic baselines. "
                "The approach adapts to feedstock variability, a persistent challenge in real-world "
                "recycling facilities."
            ),
            "source": "https://doi.org/10.3390/polym18091062",
            "impact": "medium",
        },
        {
            "category": "Generative Design",
            "title": "Polybot: Argonne's Autonomous Platform Synthesises Electrochromic Polymers",
            "description": (
                "Argonne National Laboratory's Polybot platform autonomously designs, synthesises, and "
                "characterises electrochromic polymers for smart display and energy storage applications "
                "without human intervention in the design loop. The system integrates an ML surrogate "
                "model with robotic synthesis and in-situ UV-Vis spectroscopy for closed-loop property "
                "feedback. Polybot has demonstrated precise colour-state tuning across the visible spectrum, "
                "completing design cycles in days rather than months."
            ),
            "source": "https://www.anl.gov/cnm/article/aipowered-autonomous-electronic-polymer-synthesis",
            "impact": "high",
        },
    ],
    "key_trends": [
        "Inverse design is going experimental: AI-generated polymers are now synthesised and lab-validated, not just simulated",
        "Multimodal LLMs (text + 3D structure) outperform specialist GNNs with far less labelled data, democratising polymer ML",
        "Self-driving polymer laboratories are scaling from proof-of-concept to production platforms at national labs",
        "Open large-scale datasets (OPoly26) are consolidating community benchmarking and reducing duplicated DFT effort",
        "AI-driven sorting and pyrolysis optimisation are making plastic circular-economy targets economically achievable by 2030",
    ],
    "notable_tools": [
        {
            "name": "POLYT5",
            "purpose": (
                "Encoder-decoder chemical-language foundation model for generative inverse design of polymers "
                "from target properties; first to be experimentally validated (Georgia Tech, 2026)."
            ),
            "url": "https://phys.org/news/2026-03-generative-ai-polymer-lab-dielectric.html",
        },
        {
            "name": "PolyLLMem",
            "purpose": (
                "Multimodal ML model fusing Llama 3 text and Uni-Mol structure embeddings via LoRA "
                "to predict 22 polymer properties with minimal labelled data."
            ),
            "url": "https://arxiv.org/abs/2503.22962",
        },
        {
            "name": "OPoly26 Dataset",
            "purpose": (
                "Open benchmark of 6.35 million DFT calculations on polymer systems for training and "
                "evaluating property-prediction models."
            ),
            "url": "https://arxiv.org/pdf/2512.23117",
        },
        {
            "name": "Polybot (Argonne)",
            "purpose": (
                "Autonomous robotic platform for closed-loop electrochromic polymer synthesis, "
                "characterisation, and ML-driven design iteration."
            ),
            "url": "https://www.anl.gov/cnm/article/aipowered-autonomous-electronic-polymer-synthesis",
        },
        {
            "name": "Polymer Genome",
            "purpose": (
                "Web-based informatics platform for data-driven polymer property prediction "
                "and materials screening."
            ),
            "url": "https://www.polymergenome.org",
        },
        {
            "name": "Multi-Agent Polymer Informatics",
            "purpose": (
                "Autonomous multi-agent AI system orchestrating the full pipeline from literature "
                "mining to generative design and retrosynthesis (arXiv 2602.00103)."
            ),
            "url": "https://arxiv.org/pdf/2602.00103",
        },
    ],
    "outlook": (
        "The convergence of large-scale open datasets, multimodal LLMs, and physically validated generative "
        "models signals that polymer AI is entering a phase of rapid industrial deployment rather than "
        "laboratory curiosity. Within the next 12 months we expect to see the first commercial polymer "
        "products — likely in dielectrics, electrochromics, or biodegradable packaging — that are explicitly "
        "attributed to AI-driven inverse design rather than conventional trial-and-error. Self-driving "
        "laboratories will proliferate beyond national labs into specialty chemical companies as robotics "
        "costs fall and ML toolkits mature. The sustainability imperative is a powerful tailwind: regulatory "
        "pressure under the EU PPWR and voluntary net-zero commitments are forcing polymer producers to adopt "
        "AI-optimised recycling far faster than pure R&D incentives would. The principal risk is data "
        "fragmentation — without continued community investment in open, well-curated datasets like OPoly26, "
        "model benchmarks will remain difficult to trust and the gap between well-funded and under-resourced "
        "groups will widen rather than close."
    ),
}

if __name__ == "__main__":
    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(exist_ok=True)

    docx_path = output_dir / f"Polymer_AI_Digest_{DATE}.docx"
    pptx_path = output_dir / f"Polymer_AI_Digest_{DATE}.pptx"

    print("Generating Word document...")
    generate_word(data, docx_path)

    print("Generating PowerPoint...")
    generate_pptx(data, pptx_path)

    print(f"\nFiles written to {output_dir.resolve()}")
    print(f"  {docx_path.name}")
    print(f"  {pptx_path.name}")
