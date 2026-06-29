#!/usr/bin/env python3
"""Weekly AI-for-Polymers digest generator — 2026-06-29"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from agent import generate_word, generate_pptx

DATE = "2026-06-29"

data = {
    "date": DATE,
    "headline": (
        "From lab-validated generative design to LLM polymer agents and AI-powered "
        "recycling, AI is now end-to-end across the full polymer lifecycle in mid-2026."
    ),
    "executive_summary": (
        "The final week of June 2026 consolidates a landmark first half of the year for "
        "AI in polymer materials science. Generative foundation models are no longer "
        "experimental: POLYT5 from Georgia Tech produced a novel dielectric polymer that "
        "was physically synthesized and confirmed to perform as predicted — a milestone "
        "representing the first end-to-end experimental validation of a generative polymer "
        "AI. Simultaneously, autonomous inverse-design workflows coupled with robotic "
        "experimentation narrowed more than 1000 candidate compositions to near-target "
        "recipes in under 72 hours, compressing what was once a multi-year process into "
        "days. IBM and Samsung filed major patents in 2025-2026 covering foundation-model "
        "driven substructural replacement and automated flow-chemistry feedback loops, "
        "signalling that industrial scale-up of generative polymer AI is now underway.\n\n"
        "The data infrastructure enabling these advances matured significantly with the "
        "release of OPoly26, a benchmark dataset containing more than 6.35 million DFT "
        "calculations across polymeric systems comprising over 1.2 billion atoms. Adaptive "
        "graph neural network architectures fusing GNN fingerprints with molecular "
        "representations now achieve state-of-the-art multi-task thermal property "
        "prediction, as confirmed in Macromolecules 2026. Hybrid ML frameworks for polymer "
        "solubility — spanning concentrations and temperatures — are reaching production "
        "readiness, addressing longstanding formulation challenges in coatings, adhesives, "
        "and drug delivery systems.\n\n"
        "Large language models have crossed from novelty to utility in polymer informatics. "
        "Polymer-Agent, published in JCIM 2026, provides property prediction, structure "
        "generation, and literature mining through a natural-language interface, benchmarked "
        "favorably against task-specific ML models. polyRETRO extends LLM capabilities to "
        "retrosynthetic polymer planning, while PolySea applies LLM reasoning to "
        "evolutionary sequence optimization. Benchmarking studies confirm LLMs match or "
        "exceed classical fingerprinting methods on multiple property tasks without "
        "requiring labeled training data — a transformative finding for data-sparse domains.\n\n"
        "Sustainability remains the most commercially urgent frontier. AI-driven sorting "
        "systems combining near-infrared, Raman, and hyperspectral imaging are delivering "
        "60 percent throughput gains at recycling facilities. Multi-objective AI optimization "
        "of plastic waste pyrolysis simultaneously maximizes yield and minimizes CO2 "
        "emissions. Meanwhile, deep-learning models for FTIR fingerprint prediction, "
        "physics-informed neural networks cutting cure cycle times by 30 percent, and "
        "99%-accurate defect detection systems are together establishing AI as a standard "
        "component of polymer manufacturing rather than an experimental add-on."
    ),
    "findings": [
        {
            "category": "Generative Design",
            "title": "POLYT5 Foundation Model Achieves First Lab-Validated Generative Polymer Design",
            "description": (
                "Georgia Tech researchers released POLYT5, an encoder-decoder foundation "
                "chemical language model trained on experimentally verified polymers that "
                "accepts user-specified property targets and outputs novel polymer structures. "
                "The model produced a new dielectric polymer that was synthesized and "
                "confirmed to perform as predicted in laboratory tests — the first generative "
                "polymer AI with physical experimental validation. This breakthrough "
                "demonstrates that foundation models can bridge computational prediction "
                "and real-world material performance."
            ),
            "source": "https://coe.gatech.edu/news/2026/03/researchers-create-first-ai-generative-polymer-design",
            "impact": "high",
        },
        {
            "category": "Generative Design",
            "title": "Autonomous 72-Hour Inverse Design Narrows 1000+ Compositions to Target",
            "description": (
                "An autonomous inverse-design workflow integrating AI-driven literature mining, "
                "machine learning surrogates, and robotic experimentation completed a full design "
                "cycle in under 72 hours — starting from more than a thousand candidate "
                "compositions and narrowing to near-target polymer recipes in fewer than 50 "
                "experiments. IBM (WO 2026 patent) and Samsung (2025 patent) both filed broad "
                "IP covering AI candidate generation coupled with automated flow chemistry and "
                "real-time analytical feedback, signalling industrial-scale adoption of "
                "inverse polymer design."
            ),
            "source": "https://phys.org/news/2026-03-inverse-pathway-custom-functional-polymers.html",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "OPoly26 Dataset: 6.35 Million DFT Calculations Redefine Polymer ML Benchmarking",
            "description": (
                "The Open Polymers 2026 (OPoly26) dataset provides more than 6.35 million DFT "
                "calculations on polymeric clusters of up to 360 atoms, covering over 1.2 billion "
                "total atoms across diverse monomer compositions, chain architectures, and "
                "solvation environments. Adaptive fusion of graph neural networks and molecular "
                "fingerprints via gating now achieves state-of-the-art multi-task thermal property "
                "prediction, as reported in Macromolecules 2026, establishing OPoly26 as the "
                "primary community benchmark for next-generation property prediction models."
            ),
            "source": "https://arxiv.org/pdf/2512.23117",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "Hybrid ML Frameworks Predict Polymer Solubility Across Concentrations and Temperatures",
            "description": (
                "New hybrid machine learning frameworks combining polymer composition and sequence "
                "features predict solubility across wide ranges of concentrations and temperatures "
                "with production-ready accuracy, addressing a longstanding formulation challenge in "
                "coatings, adhesives, and pharmaceutical delivery. White-box ML surrogates for "
                "charged polymer blend phase prediction deliver fast, interpretable outputs matching "
                "computationally expensive simulation methods, enabling real-time process decisions "
                "during manufacturing scale-up."
            ),
            "source": "https://www.ncbi.nlm.nih.gov/pmc/articles/PMC11684024/",
            "impact": "medium",
        },
        {
            "category": "LLMs",
            "title": "Polymer-Agent: LLM Agent for End-to-End Polymer Design Goes Live in JCIM",
            "description": (
                "Polymer-Agent, published in the Journal of Chemical Information and Modeling "
                "in 2026, provides property prediction, property-guided polymer structure "
                "generation, and structure modification through a conversational natural-language "
                "interface powered by LLM reasoning. Benchmarking studies confirm LLM-based "
                "methods match or exceed conventional ML approaches on several polymer property "
                "tasks without requiring complex feature engineering or large labeled datasets, "
                "lowering the barrier for domain experts who lack ML expertise."
            ),
            "source": "https://pubs.acs.org/doi/10.1021/acs.jcim.6c00343",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "polyRETRO and PolySea Extend LLMs to Retrosynthesis and Sequence Optimization",
            "description": (
                "polyRETRO, an LLM-based model for predicting polymerization class and required "
                "monomers for a target polymer, enables retrosynthetic planning at the polymer "
                "level — previously impossible with standard chemistry retrosynthesis tools. "
                "PolySea (Polymer Smart Evolution Agent) applies LLM reasoning loops to "
                "evolutionary sequence optimization, generating diverse candidate backbones "
                "and iterating toward property targets autonomously. Together they establish "
                "a growing LLM toolkit for both forward synthesis planning and inverse design."
            ),
            "source": "https://arxiv.org/pdf/2512.05138",
            "impact": "medium",
        },
        {
            "category": "Characterization",
            "title": "Deep Learning Achieves 87% FTIR Fingerprint Prediction and Sub-10 µm Microplastic Detection",
            "description": (
                "A 2026 MRS Advances study optimized multilayer perceptron architectures for FTIR "
                "fingerprint-region prediction, reaching 87% accuracy for polystyrene, 84% for "
                "HDPE, and 81% for PET among six commodity plastics — substantially exceeding "
                "prior reference-library matching. Separately, hyperspectral deep learning now "
                "achieves high-accuracy segmentation of sub-10 µm microplastics, enabling "
                "automated environmental monitoring at particle sizes previously requiring "
                "labor-intensive manual analysis under electron microscopy."
            ),
            "source": "https://link.springer.com/article/10.1557/s43580-026-01613-8",
            "impact": "medium",
        },
        {
            "category": "Processing",
            "title": "Physics-Informed Neural Networks Cut Cure Times 30% and Defect Rates Below 1%",
            "description": (
                "Physics-informed neural networks applied to composite autoclave processing predict "
                "optimal heating curves that reduce cure cycle times by up to 30% without "
                "sacrificing part quality, as documented in a 2026 MRS Communications review. "
                "AI-driven predictive quality systems in additive manufacturing now exceed 99% "
                "defect detection accuracy, with computer vision monitoring every layer during "
                "SLS and SLA printing in real time. Real-time ML process control simultaneously "
                "reduces off-spec production by over 2% and natural-gas consumption by 10-20%."
            ),
            "source": "https://link.springer.com/article/10.1557/s43579-026-00972-5",
            "impact": "high",
        },
        {
            "category": "Sustainability",
            "title": "AI-Driven Sorting Delivers 60% Efficiency Gains; Pyrolysis Optimization Scales",
            "description": (
                "Deep-learning sorting systems combining near-infrared, Raman, and hyperspectral "
                "imaging are routing plastics to optimal recycling pathways with up to 60% higher "
                "throughput than manual sorting, per 2026 industry data from Plastics Today. "
                "Multi-objective AI optimization of plastic waste pyrolysis (Polymers 18(9):1062) "
                "simultaneously maximizes oil yield and minimizes energy input and CO2 emissions, "
                "making thermochemical recycling of mixed plastic streams economically viable "
                "at industrial scale for the first time."
            ),
            "source": "https://doi.org/10.3390/polym18091062",
            "impact": "high",
        },
        {
            "category": "Sustainability",
            "title": "AI-Guided Discovery of Recyclable Vitrimeric Polymers Advances Circular Thermosets",
            "description": (
                "AI-guided inverse design workflows are being applied to vitrimer chemistry, "
                "enabling discovery of recyclable thermoset polymers with targeted mechanical "
                "and thermal performance. By coupling ML property prediction with sustainability "
                "constraints — recyclability index, bio-based feedstock compatibility, and "
                "depolymerization pathway scoring — researchers are identifying vitrimer "
                "formulations that meet both performance and end-of-life requirements "
                "simultaneously, advancing circular economy goals for the thermoset market."
            ),
            "source": "https://arxiv.org/pdf/2312.03690",
            "impact": "medium",
        },
    ],
    "key_trends": [
        "Lab-validated generative AI: inverse-design models confirmed by physical synthesis experiments",
        "LLM agents democratizing polymer design — natural language removes the ML expertise barrier",
        "OPoly26 and OpenPoly mega-datasets enabling rigorous, reproducible ML benchmarking",
        "AI closing the recycling loop: spectroscopy + deep learning delivering 60%+ sorting gains",
        "Physics-informed NNs bridging process simulation and ML for 30% manufacturing cycle reductions",
    ],
    "notable_tools": [
        {
            "name": "POLYT5",
            "purpose": (
                "Foundation encoder-decoder language model for generative polymer design; "
                "accepts property specs and outputs lab-validated novel structures."
            ),
            "url": "https://coe.gatech.edu/news/2026/03/researchers-create-first-ai-generative-polymer-design",
        },
        {
            "name": "Polymer-Agent (JCIM 2026)",
            "purpose": (
                "LLM-based conversational agent providing property prediction, structure "
                "generation, and literature mining for polymer scientists without ML expertise."
            ),
            "url": "https://pubs.acs.org/doi/10.1021/acs.jcim.6c00343",
        },
        {
            "name": "OPoly26 Dataset",
            "purpose": (
                "Open benchmark with 6.35M DFT calculations across polymer systems; "
                "the primary community resource for next-gen property prediction ML."
            ),
            "url": "https://arxiv.org/pdf/2512.23117",
        },
        {
            "name": "polyRETRO",
            "purpose": (
                "LLM approach to retrosynthetic polymer planning: predicts polymerization "
                "class and required monomers for a user-specified target polymer."
            ),
            "url": "https://arxiv.org/pdf/2512.05138",
        },
        {
            "name": "OpenPoly",
            "purpose": (
                "Curated polymer database empowering ML benchmarking and multi-property "
                "predictions; published in Chinese Journal of Polymer Science, 2025."
            ),
            "url": "https://link.springer.com/article/10.1007/s10118-025-3402-y",
        },
        {
            "name": "PolyInfo (NIMS PoLyInfo)",
            "purpose": (
                "Foundational polymer property database with 500K+ experimental data points "
                "covering thermal, electrical, and mechanical properties across 100+ property types."
            ),
            "url": "https://polymer.nims.go.jp/",
        },
    ],
    "outlook": (
        "The trajectory for AI in polymer materials science points unmistakably toward "
        "full-stack autonomy within the next 12-18 months. Self-driving laboratories "
        "combining generative AI, robotic synthesis, and real-time characterization are "
        "expected to compress targeted polymer development from years to weeks, and early "
        "prototypes are already operational at leading academic centers. LLM agents will "
        "evolve from single-task assistants to orchestrators managing the entire workflow — "
        "from literature mining through synthesis planning to process optimization — making "
        "advanced polymer design accessible to domain scientists without ML backgrounds. "
        "The central bottleneck is shifting from algorithmic capability to data governance: "
        "standardized, FAIR-compliant polymer property databases with full experimental "
        "provenance will determine which organizations can train trustworthy, generalizable "
        "models. Sustainability remains the dominant commercial driver; AI-optimized chemical "
        "recycling, biodegradability-by-design, and closed-loop circular manufacturing are "
        "receiving disproportionate investment as regulatory pressure on plastic waste and "
        "carbon emissions intensifies globally through 2027."
    ),
}

if __name__ == "__main__":
    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(exist_ok=True)

    docx_path = output_dir / f"Polymer_AI_Digest_{DATE}.docx"
    pptx_path = output_dir / f"Polymer_AI_Digest_{DATE}.pptx"

    print(f"Generating Word document -> {docx_path}")
    generate_word(data, docx_path)

    print(f"Generating PowerPoint   -> {pptx_path}")
    generate_pptx(data, pptx_path)

    print("\nDone.")
